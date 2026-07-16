"""
FastAPI application for SharePoint site management aligned with
CRM/Deal Center business requirements.

Business flows:
1. NEW COMPANY -> Create subsite + folders + upload docs
2. EXISTING COMPANY (no subsite) + Opportunity -> Create subsite + folders + upload docs
3. EXISTING COMPANY (has subsite) + Opportunity -> Create folders + upload docs only
4. USER UPLOAD -> Upload docs to existing folders

Endpoints:
- GET /sites - List SharePoint sites
- GET /sites/{site_id}/subsites - List subsites
- POST /sites/{site_id}/subsites - Create subsite
- DRIVE: get drive, list children, create folder, upload file, download, delete, search
- WORKFLOWS:
  - POST /workflow/company-onboarding - Full company setup (subsite + folders + docs)
  - POST /workflow/opportunity-processing - Opportunity-based folder/docs creation
  - POST /workflow/user-upload - Upload documents to existing folders
  - POST /workflow/subsite-lookup - Check if subsite exists

Environment variables:
- AZURE_CLIENT_ID, AZURE_TENANT_ID, AZURE_CLIENT_SECRET (or Managed Identity)
- SHAREPOINT_HOSTNAME
- DEFAULT_FOLDER_TEMPLATE (optional JSON array of folder names)
"""

from __future__ import annotations

import asyncio
import base64
import io
import json
import os
import time
from typing import Any, Dict, List, Optional
from urllib.parse import urlparse

import httpx
from azure.identity import DefaultAzureCredential
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic import BaseModel, Field

app = FastAPI(
    title="SharePoint CRM Integration API",
    description="Backend API for Deal Center CRM -> SharePoint integration",
    version="5.0.0",
)

# ------------------------------------------------------------------
# Configuration
# ------------------------------------------------------------------

GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0"
SHAREPOINT_HOSTNAME = os.getenv("SHAREPOINT_HOSTNAME", "yourtenant.sharepoint.com")

DEFAULT_FOLDER_TEMPLATE = os.getenv(
    "DEFAULT_FOLDER_TEMPLATE",
    json.dumps(["Contracts", "Proposals", "Invoices", "Legal", "Communications"]),
)

credential = DefaultAzureCredential()

# ------------------------------------------------------------------
# Pydantic models
# ------------------------------------------------------------------

class SiteResponse(BaseModel):
    id: str
    name: str
    webUrl: str
    displayName: Optional[str] = None
    description: Optional[str] = None
    createdDateTime: Optional[str] = None
    lastModifiedDateTime: Optional[str] = None

class SubsiteCreateRequest(BaseModel):
    title: str = Field(..., min_length=1, max_length=255)
    description: Optional[str] = Field(None, max_length=500)
    url: str = Field(..., min_length=1, max_length=128)
    template: str = Field("STS#3")
    language: int = Field(1033)
    use_same_permissions: bool = Field(False)

class SubsiteResponse(BaseModel):
    id: str
    title: str
    url: str
    serverRelativeUrl: str
    created: Optional[str] = None
    webTemplate: Optional[str] = None
    webUrl: Optional[str] = None
    graphSiteId: Optional[str] = None

class DriveResponse(BaseModel):
    id: str
    name: str
    driveType: str
    webUrl: Optional[str] = None
    owner: Optional[Dict[str, Any]] = None
    quota: Optional[Dict[str, Any]] = None

class DriveItemResponse(BaseModel):
    id: str
    name: str
    size: Optional[int] = None
    webUrl: Optional[str] = None
    createdDateTime: Optional[str] = None
    lastModifiedDateTime: Optional[str] = None
    folder: Optional[Dict[str, Any]] = None
    file: Optional[Dict[str, Any]] = None
    parentReference: Optional[Dict[str, Any]] = None
    microsoft_graph_conflictBehavior: Optional[str] = Field(None, alias="@microsoft.graph.conflictBehavior")

    class Config:
        populate_by_name = True

class FolderCreateRequest(BaseModel):
    name: str = Field(..., min_length=1, max_length=255)
    conflict_behavior: str = Field("rename")

class FileUploadResponse(BaseModel):
    id: str
    name: str
    size: int
    webUrl: str
    createdDateTime: str
    lastModifiedDateTime: str
    downloadUrl: Optional[str] = None

class FolderContentsResponse(BaseModel):
    value: List[DriveItemResponse]
    count: int
    nextLink: Optional[str] = None

# -- Workflow Models --

class CompanyOnboardingRequest(BaseModel):
    parent_site_id: str = Field(..., description="Parent site ID (e.g., the main deal site)")
    company_name: str = Field(..., min_length=1, max_length=255)
    company_url_slug: str = Field(..., min_length=1, max_length=128, description="URL-friendly name, e.g., 'acme-corp'")
    description: Optional[str] = Field(None, description="Company description")
    folders: Optional[List[str]] = Field(None, description="Custom folder structure. If None, uses DEFAULT_FOLDER_TEMPLATE")
    initial_documents: Optional[List[Dict[str, str]]] = Field(None, description="List of {filename, content_base64} for initial uploads")

class CompanyOnboardingResponse(BaseModel):
    company_name: str
    subsite: SubsiteResponse
    drive: Optional[DriveResponse] = None
    created_folders: List[DriveItemResponse]
    uploaded_documents: List[FileUploadResponse]
    message: str

class OpportunityProcessingRequest(BaseModel):
    subsite_id: str = Field(..., description="Existing subsite Graph ID")
    opportunity_name: str = Field(..., min_length=1, max_length=255)
    opportunity_folder_name: Optional[str] = Field(None, description="Folder name for this opportunity. Defaults to opportunity_name")
    documents: Optional[List[Dict[str, str]]] = Field(None, description="Documents to upload: {filename, content_base64}")
    create_opportunity_folder: bool = Field(True, description="Create a dedicated opportunity folder")

class OpportunityProcessingResponse(BaseModel):
    opportunity_name: str
    subsite_id: str
    opportunity_folder: Optional[DriveItemResponse] = None
    created_folders: List[DriveItemResponse]
    uploaded_documents: List[FileUploadResponse]
    message: str

class UserUploadRequest(BaseModel):
    site_id: str = Field(..., description="Site or subsite Graph ID")
    folder_path: str = Field(..., description="Folder path or item ID. Use 'root' for root, or folder ID")
    files: List[Dict[str, str]] = Field(..., description="List of {filename, content_base64} to upload")

class UserUploadResponse(BaseModel):
    site_id: str
    folder_path: str
    uploaded_files: List[FileUploadResponse]
    failed_files: List[Dict[str, str]]
    message: str

class SubsiteLookupRequest(BaseModel):
    parent_site_id: str
    company_url_slug: str

class SubsiteLookupResponse(BaseModel):
    exists: bool
    subsite: Optional[SiteResponse] = None
    message: str

# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def _parse_site_web_url(site_web_url: str) -> tuple[str, str]:
    parsed = urlparse(site_web_url)
    if not parsed.scheme or not parsed.netloc:
        raise ValueError(f"Invalid SharePoint webUrl: {site_web_url}")
    return parsed.netloc, parsed.path.rstrip("/")

async def get_graph_token() -> str:
    try:
        token = credential.get_token("https://graph.microsoft.com/.default")
        return token.token
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to acquire Graph API token: {exc}") from exc

async def get_sharepoint_token() -> str:
    try:
        token = credential.get_token(f"https://{SHAREPOINT_HOSTNAME}/.default")
        return token.token
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to acquire SharePoint token: {exc}") from exc

async def make_graph_request(
    method: str,
    endpoint: str,
    json_data: Optional[Dict[str, Any]] = None,
    content: Optional[bytes] = None,
    content_type: Optional[str] = None,
    max_retries: int = 3,
) -> Dict[str, Any]:
    token = await get_graph_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }
    if json_data is not None:
        headers["Content-Type"] = "application/json"
    if content_type:
        headers["Content-Type"] = content_type

    url = f"{GRAPH_BASE_URL}{endpoint}"

    async with httpx.AsyncClient(timeout=60.0) as client:
        for attempt in range(max_retries):
            try:
                method_upper = method.upper()
                if method_upper == "GET":
                    response = await client.get(url, headers=headers)
                elif method_upper == "POST":
                    response = await client.post(url, headers=headers, content=content) if content is not None else await client.post(url, headers=headers, json=json_data)
                elif method_upper == "PUT":
                    response = await client.put(url, headers=headers, content=content) if content is not None else await client.put(url, headers=headers, json=json_data)
                elif method_upper == "DELETE":
                    response = await client.delete(url, headers=headers)
                elif method_upper == "PATCH":
                    response = await client.patch(url, headers=headers, json=json_data)
                else:
                    raise ValueError(f"Unsupported HTTP method: {method}")

                if response.status_code == 204:
                    return {}
                response.raise_for_status()
                return response.json()

            except httpx.HTTPStatusError as exc:
                if exc.response.status_code == 429 and attempt < max_retries - 1:
                    retry_after = int(exc.response.headers.get("Retry-After", "2"))
                    await asyncio.sleep(retry_after)
                    continue

                error_detail = exc.response.text
                error_code = "HTTPError"
                try:
                    error_json = exc.response.json()
                    error_detail = error_json.get("error", {}).get("message", error_detail)
                    error_code = error_json.get("error", {}).get("code", error_code)
                except Exception:
                    pass
                raise HTTPException(
                    status_code=exc.response.status_code,
                    detail=f"Graph API Error ({error_code}): {error_detail}",
                ) from exc
            except Exception as exc:
                raise HTTPException(status_code=500, detail=f"Request failed: {exc}") from exc

        raise HTTPException(status_code=500, detail="Graph request failed after retries")

async def make_sharepoint_request(
    site_web_url: str,
    endpoint: str,
    json_data: Optional[Dict[str, Any]] = None,
    method: str = "POST",
) -> Dict[str, Any]:
    token = await get_sharepoint_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json;odata=verbose",
        "Accept": "application/json;odata=verbose",
    }

    parsed = urlparse(site_web_url)
    if not parsed.scheme or not parsed.netloc:
        raise HTTPException(status_code=400, detail=f"Invalid SharePoint site URL: {site_web_url}")

    url = f"{parsed.scheme}://{parsed.netloc}{parsed.path.rstrip('/')}{endpoint}"

    async with httpx.AsyncClient(timeout=60.0) as client:
        try:
            method_upper = method.upper()
            if method_upper == "GET":
                response = await client.get(url, headers=headers)
            elif method_upper == "POST":
                response = await client.post(url, headers=headers, json=json_data)
            else:
                raise ValueError(f"Unsupported HTTP method: {method}")

            response.raise_for_status()
            data = response.json()
            return data.get("d", data)

        except httpx.HTTPStatusError as exc:
            error_detail = exc.response.text
            try:
                error_json = exc.response.json()
                error_detail = error_json.get("error", {}).get("message", {}).get("value", error_detail)
            except Exception:
                pass
            raise HTTPException(status_code=exc.response.status_code, detail=f"SharePoint API Error: {error_detail}") from exc
        except Exception as exc:
            raise HTTPException(status_code=500, detail=f"SharePoint request failed: {exc}") from exc

async def find_subsite_graph_id(parent_web_url: str, subsite_url_segment: str, title: str, max_wait: int = 30) -> Optional[str]:
    expected_url = f"{parent_web_url.rstrip('/')}/{subsite_url_segment}"
    start_time = time.time()
    while time.time() - start_time < max_wait:
        try:
            # FIXED: Proper search syntax with double quotes
            result = await make_graph_request("GET", f'/sites?search="{title}"&$top=10')
            for site in result.get("value", []):
                if site.get("webUrl", "").rstrip("/") == expected_url.rstrip("/"):
                    return site.get("id")

            # Fallback: direct lookup by hostname + path
            host = expected_url.replace("https://", "").split("/")[0]
            path = "/".join(expected_url.replace("https://", "").split("/")[1:])
            try:
                direct_result = await make_graph_request("GET", f"/sites/{host}:/{path}")
                if direct_result.get("id"):
                    return direct_result.get("id")
            except Exception:
                pass
        except HTTPException:
            pass
        await asyncio.sleep(2)
    return None

async def resolve_graph_site_id(parent_web_url: str, subsite_url_segment: str, title: str, max_wait: int = 30) -> str:
    graph_site_id = await find_subsite_graph_id(parent_web_url, subsite_url_segment, title, max_wait=max_wait)
    if graph_site_id:
        return graph_site_id

    # Fallback: construct path-based ID
    host = parent_web_url.replace("https://", "").split("/")[0]
    path = "/".join(f"{parent_web_url.rstrip('/')}/{subsite_url_segment}".replace("https://", "").split("/")[1:])

    try:
        lookup = await make_graph_request("GET", f"/sites/{host}:/{path}")
        if lookup.get("id"):
            return lookup["id"]
    except Exception:
        pass

    # Final fallback: return path-based ID anyway — Graph API might accept it
    return f"{host}:/{path}"

# -- Internal Drive Helpers --

async def _get_site_drive(site_id: str) -> DriveResponse:
    result = await make_graph_request("GET", f"/sites/{site_id}/drive")
    return DriveResponse(
        id=result.get("id", ""),
        name=result.get("name", ""),
        driveType=result.get("driveType", "documentLibrary"),
        webUrl=result.get("webUrl"),
        owner=result.get("owner"),
        quota=result.get("quota"),
    )

async def _create_folder(site_id: str, parent_item_id: str, folder_name: str, conflict_behavior: str = "rename") -> DriveItemResponse:
    payload = {
        "name": folder_name,
        "folder": {},
        "@microsoft.graph.conflictBehavior": conflict_behavior,
    }
    result = await make_graph_request("POST", f"/sites/{site_id}/drive/items/{parent_item_id}/children", payload)
    return DriveItemResponse(
        id=result.get("id", ""),
        name=result.get("name", ""),
        size=result.get("size"),
        webUrl=result.get("webUrl"),
        createdDateTime=result.get("createdDateTime"),
        lastModifiedDateTime=result.get("lastModifiedDateTime"),
        folder=result.get("folder"),
        file=result.get("file"),
        parentReference=result.get("parentReference"),
        microsoft_graph_conflictBehavior=result.get("@microsoft.graph.conflictBehavior"),
    )

async def _upload_file_from_bytes(
    site_id: str,
    parent_item_id: str,
    filename: str,
    content: bytes,
    content_type: str = "application/octet-stream",
) -> FileUploadResponse:
    result = await make_graph_request(
        "PUT",
        f"/sites/{site_id}/drive/items/{parent_item_id}:/{filename}:/content",
        content=content,
        content_type=content_type,
    )
    return FileUploadResponse(
        id=result.get("id", ""),
        name=result.get("name", filename),
        size=result.get("size", len(content)),
        webUrl=result.get("webUrl", ""),
        createdDateTime=result.get("createdDateTime", ""),
        lastModifiedDateTime=result.get("lastModifiedDateTime", ""),
        downloadUrl=result.get("@microsoft.graph.downloadUrl"),
    )

async def _create_subsite_internal(site_web_url: str, request: SubsiteCreateRequest) -> SubsiteResponse:
    payload = {
        "parameters": {
            "__metadata": {"type": "SP.WebInfoCreationInformation"},
            "Title": request.title,
            "Url": request.url,
            "Description": request.description or "",
            "Language": request.language,
            "WebTemplate": request.template,
            "UseSamePermissionsAsParentSite": request.use_same_permissions,
        }
    }
    result = await make_sharepoint_request(site_web_url, "/_api/web/webinfos/add", payload)
    web_info = result.get("d", result)
    subsite_web_url = f"{site_web_url.rstrip('/')}/{request.url}"

    return SubsiteResponse(
        id=web_info.get("Id", web_info.get("ID", "")),
        title=web_info.get("Title", request.title),
        url=web_info.get("Url", request.url),
        serverRelativeUrl=web_info.get("ServerRelativeUrl", ""),
        created=web_info.get("Created"),
        webTemplate=web_info.get("WebTemplate", request.template),
        webUrl=subsite_web_url,
    )

# ------------------------------------------------------------------
# Routes
# ------------------------------------------------------------------

@app.get("/")
async def root() -> Dict[str, Any]:
    return {
        "service": "SharePoint CRM Integration API",
        "status": "running",
        "auth_method": "Azure Identity (DefaultAzureCredential)",
        "flows": {
            "1_company_onboarding": "POST /workflow/company-onboarding",
            "2_opportunity_processing": "POST /workflow/opportunity-processing",
            "3_user_upload": "POST /workflow/user-upload",
            "4_subsite_lookup": "POST /workflow/subsite-lookup",
        },
    }

@app.get("/sites", response_model=List[SiteResponse])
async def list_sites(search: Optional[str] = None, top: int = 100):
    try:
        params = f"?$top={min(top, 999)}"
        if search:
            # FIXED: Proper search syntax
            params += f'&search="{search}"'
        result = await make_graph_request("GET", f"/sites{params}")
        sites = result.get("value", [])
        return [
            SiteResponse(
                id=s.get("id", ""),
                name=s.get("name", ""),
                webUrl=s.get("webUrl", ""),
                displayName=s.get("displayName"),
                description=s.get("description"),
                createdDateTime=s.get("createdDateTime"),
                lastModifiedDateTime=s.get("lastModifiedDateTime"),
            )
            for s in sites
        ]
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to list sites: {exc}") from exc

@app.get("/sites/{site_id}", response_model=SiteResponse)
async def get_site(site_id: str):
    try:
        result = await make_graph_request("GET", f"/sites/{site_id}")
        return SiteResponse(
            id=result.get("id", ""),
            name=result.get("name", ""),
            webUrl=result.get("webUrl", ""),
            displayName=result.get("displayName"),
            description=result.get("description"),
            createdDateTime=result.get("createdDateTime"),
            lastModifiedDateTime=result.get("lastModifiedDateTime"),
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to get site: {exc}") from exc

@app.get("/sites/{site_id}/subsites", response_model=List[SiteResponse])
async def list_subsites(site_id: str, top: int = 100):
    try:
        params = f"?$top={min(top, 999)}"
        result = await make_graph_request("GET", f"/sites/{site_id}/sites{params}")
        subsites = result.get("value", [])
        return [
            SiteResponse(
                id=s.get("id", ""),
                name=s.get("name", ""),
                webUrl=s.get("webUrl", ""),
                displayName=s.get("displayName"),
                description=s.get("description"),
                createdDateTime=s.get("createdDateTime"),
                lastModifiedDateTime=s.get("lastModifiedDateTime"),
            )
            for s in subsites
        ]
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to list subsites: {exc}") from exc

@app.post("/sites/{site_id}/subsites", response_model=SubsiteResponse, status_code=201)
async def create_subsite(site_id: str, request: SubsiteCreateRequest):
    try:
        parent = await get_site(site_id)
        return await _create_subsite_internal(parent.webUrl, request)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to create subsite: {exc}") from exc

# -- DRIVE / FOLDER / FILE ENDPOINTS --

@app.get("/sites/{site_id}/drive", response_model=DriveResponse)
async def get_site_drive(site_id: str):
    try:
        return await _get_site_drive(site_id)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to get drive: {exc}") from exc

@app.get("/sites/{site_id}/drive/items/{item_id}/children", response_model=FolderContentsResponse)
async def list_drive_item_children(site_id: str, item_id: str, top: int = 200, select: Optional[str] = None):
    try:
        params = f"?$top={min(top, 999)}"
        if select:
            params += f"&$select={select}"
        result = await make_graph_request("GET", f"/sites/{site_id}/drive/items/{item_id}/children{params}")
        items = result.get("value", [])
        return FolderContentsResponse(
            value=[
                DriveItemResponse(
                    id=i.get("id", ""),
                    name=i.get("name", ""),
                    size=i.get("size"),
                    webUrl=i.get("webUrl"),
                    createdDateTime=i.get("createdDateTime"),
                    lastModifiedDateTime=i.get("lastModifiedDateTime"),
                    folder=i.get("folder"),
                    file=i.get("file"),
                    parentReference=i.get("parentReference"),
                    microsoft_graph_conflictBehavior=i.get("@microsoft.graph.conflictBehavior"),
                )
                for i in items
            ],
            count=len(items),
            nextLink=result.get("@odata.nextLink"),
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to list folder contents: {exc}") from exc

@app.get("/sites/{site_id}/drive/root/children", response_model=FolderContentsResponse)
async def list_drive_root_children(site_id: str, top: int = 200):
    return await list_drive_item_children(site_id, "root", top)

@app.post("/sites/{site_id}/drive/items/{item_id}/children", response_model=DriveItemResponse, status_code=201)
async def create_folder(site_id: str, item_id: str, request: FolderCreateRequest):
    try:
        return await _create_folder(site_id, item_id, request.name, request.conflict_behavior)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to create folder: {exc}") from exc

@app.put("/sites/{site_id}/drive/items/{item_id}:/{filename}:/content", response_model=FileUploadResponse)
async def upload_file(site_id: str, item_id: str, filename: str, file: UploadFile = File(...)):
    try:
        content = await file.read()
        content_type = file.content_type or "application/octet-stream"
        return await _upload_file_from_bytes(site_id, item_id, filename, content, content_type)
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to upload file: {exc}") from exc

@app.get("/sites/{site_id}/drive/items/{item_id}", response_model=DriveItemResponse)
async def get_drive_item(site_id: str, item_id: str):
    try:
        result = await make_graph_request("GET", f"/sites/{site_id}/drive/items/{item_id}")
        return DriveItemResponse(
            id=result.get("id", ""),
            name=result.get("name", ""),
            size=result.get("size"),
            webUrl=result.get("webUrl"),
            createdDateTime=result.get("createdDateTime"),
            lastModifiedDateTime=result.get("lastModifiedDateTime"),
            folder=result.get("folder"),
            file=result.get("file"),
            parentReference=result.get("parentReference"),
            microsoft_graph_conflictBehavior=result.get("@microsoft.graph.conflictBehavior"),
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to get drive item: {exc}") from exc

@app.delete("/sites/{site_id}/drive/items/{item_id}")
async def delete_drive_item(site_id: str, item_id: str):
    try:
        await make_graph_request("DELETE", f"/sites/{site_id}/drive/items/{item_id}")
        return {"message": "Item deleted successfully", "item_id": item_id}
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to delete item: {exc}") from exc

@app.get("/sites/{site_id}/drive/items/{item_id}/content")
async def download_file(site_id: str, item_id: str):
    try:
        token = await get_graph_token()
        headers = {"Authorization": f"Bearer {token}"}
        meta = await make_graph_request("GET", f"/sites/{site_id}/drive/items/{item_id}?$select=id,name,file,@microsoft.graph.downloadUrl")
        download_url = meta.get("@microsoft.graph.downloadUrl")
        filename = meta.get("name", "download")
        if not download_url:
            raise HTTPException(status_code=404, detail="File does not have a download URL")

        async with httpx.AsyncClient() as client:
            response = await client.get(download_url, headers=headers, follow_redirects=True)
            response.raise_for_status()
            content_type = response.headers.get("content-type", "application/octet-stream")
            return StreamingResponse(
                io.BytesIO(response.content),
                media_type=content_type,
                headers={"Content-Disposition": f'attachment; filename="{filename}"'},
            )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to download file: {exc}") from exc

@app.get("/sites/{site_id}/drive/search")
async def search_drive_items(site_id: str, q: str, top: int = 25):
    try:
        result = await make_graph_request("GET", f"/sites/{site_id}/drive/search(q='{q}')?$top={min(top, 999)}")
        items = result.get("value", [])
        return {
            "value": [
                {
                    "id": i.get("id"),
                    "name": i.get("name"),
                    "webUrl": i.get("webUrl"),
                    "size": i.get("size"),
                    "folder": i.get("folder"),
                    "file": i.get("file"),
                    "createdDateTime": i.get("createdDateTime"),
                    "lastModifiedDateTime": i.get("lastModifiedDateTime"),
                }
                for i in items
            ],
            "count": len(items),
            "nextLink": result.get("@odata.nextLink"),
        }
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Failed to search drive: {exc}") from exc

# ------------------------------------------------------------------
# Workflows
# ------------------------------------------------------------------

@app.post("/workflow/company-onboarding", response_model=CompanyOnboardingResponse, status_code=201)
async def company_onboarding(request: CompanyOnboardingRequest):
    """Flow 1 & 2: complete company onboarding."""
    try:
        # Step 1: Create subsite via SharePoint REST API
        parent_site = await get_site(request.parent_site_id)
        subsite_req = SubsiteCreateRequest(
            title=request.company_name,
            description=request.description,
            url=request.company_url_slug,
            template="STS#3",
            language=1033,
            use_same_permissions=False,
        )
        subsite = await _create_subsite_internal(parent_site.webUrl, subsite_req)

        # Step 2: Resolve the actual Graph site ID for the newly-created subsite
        graph_site_id = await resolve_graph_site_id(parent_site.webUrl, request.company_url_slug, request.company_name, max_wait=30)
        subsite.graphSiteId = graph_site_id

        # Step 3: Get drive and create folders
        drive = await _get_site_drive(graph_site_id)
        folder_names = request.folders or json.loads(DEFAULT_FOLDER_TEMPLATE)
        created_folders: List[DriveItemResponse] = []
        for folder_name in folder_names:
            try:
                folder = await _create_folder(graph_site_id, "root", folder_name, "rename")
                created_folders.append(folder)
            except Exception:
                pass

        # Step 4: Upload initial documents
        uploaded_documents: List[FileUploadResponse] = []
        if request.initial_documents:
            for doc in request.initial_documents:
                try:
                    filename = doc.get("filename")
                    content_b64 = doc.get("content_base64")
                    if filename and content_b64:
                        content = base64.b64decode(content_b64)
                        uploaded = await _upload_file_from_bytes(graph_site_id, "root", filename, content)
                        uploaded_documents.append(uploaded)
                except Exception:
                    pass

        return CompanyOnboardingResponse(
            company_name=request.company_name,
            subsite=subsite,
            drive=drive,
            created_folders=created_folders,
            uploaded_documents=uploaded_documents,
            message=f"Company '{request.company_name}' onboarded successfully. Subsite: {subsite.webUrl}",
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Company onboarding failed: {exc}") from exc

@app.post("/workflow/opportunity-processing", response_model=OpportunityProcessingResponse, status_code=201)
async def opportunity_processing(request: OpportunityProcessingRequest):
    """Flow 3: process an opportunity for an existing company with subsite mapping."""
    try:
        _ = await _get_site_drive(request.subsite_id)
        parent_folder_id = "root"
        opportunity_folder = None
        folder_name = request.opportunity_folder_name or request.opportunity_name

        if request.create_opportunity_folder:
            try:
                opportunity_folder = await _create_folder(request.subsite_id, "root", folder_name, "rename")
                parent_folder_id = opportunity_folder.id
            except Exception:
                pass

        created_folders: List[DriveItemResponse] = []
        for subfolder in ["Documents", "Contracts", "Communications"]:
            try:
                folder = await _create_folder(request.subsite_id, parent_folder_id, subfolder, "rename")
                created_folders.append(folder)
            except Exception:
                pass

        uploaded_documents: List[FileUploadResponse] = []
        if request.documents:
            for doc in request.documents:
                try:
                    filename = doc.get("filename")
                    content_b64 = doc.get("content_base64")
                    if filename and content_b64:
                        content = base64.b64decode(content_b64)
                        uploaded = await _upload_file_from_bytes(request.subsite_id, parent_folder_id, filename, content)
                        uploaded_documents.append(uploaded)
                except Exception:
                    pass

        return OpportunityProcessingResponse(
            opportunity_name=request.opportunity_name,
            subsite_id=request.subsite_id,
            opportunity_folder=opportunity_folder,
            created_folders=created_folders,
            uploaded_documents=uploaded_documents,
            message=f"Opportunity '{request.opportunity_name}' processed. Documents uploaded to subsite.",
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Opportunity processing failed: {exc}") from exc

@app.post("/workflow/user-upload", response_model=UserUploadResponse)
async def user_upload(request: UserUploadRequest):
    """Flow 4 & 5: upload documents to existing folders."""
    try:
        uploaded_files: List[FileUploadResponse] = []
        failed_files: List[Dict[str, str]] = []
        for file_info in request.files:
            try:
                filename = file_info.get("filename")
                content_b64 = file_info.get("content_base64")
                if not filename or not content_b64:
                    failed_files.append({"filename": filename or "unknown", "reason": "Missing filename or content"})
                    continue
                content = base64.b64decode(content_b64)
                uploaded = await _upload_file_from_bytes(request.site_id, request.folder_path, filename, content)
                uploaded_files.append(uploaded)
            except Exception as exc:
                failed_files.append({"filename": file_info.get("filename", "unknown"), "reason": str(exc)})

        return UserUploadResponse(
            site_id=request.site_id,
            folder_path=request.folder_path,
            uploaded_files=uploaded_files,
            failed_files=failed_files,
            message=f"Uploaded {len(uploaded_files)} files. Failed: {len(failed_files)}",
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"User upload failed: {exc}") from exc

@app.post("/workflow/subsite-lookup", response_model=SubsiteLookupResponse)
async def subsite_lookup(request: SubsiteLookupRequest):
    """Check if a company already has a subsite under a parent site."""
    try:
        parent_site = await get_site(request.parent_site_id)
        expected_url = f"{parent_site.webUrl.rstrip('/')}/{request.company_url_slug}"
        result = await make_graph_request("GET", f"/sites/{request.parent_site_id}/sites?$top=999")
        subsites = result.get("value", [])
        for sub in subsites:
            if sub.get("webUrl", "").rstrip("/") == expected_url.rstrip("/"):
                return SubsiteLookupResponse(
                    exists=True,
                    subsite=SiteResponse(
                        id=sub.get("id", ""),
                        name=sub.get("name", ""),
                        webUrl=sub.get("webUrl", ""),
                        displayName=sub.get("displayName"),
                        description=sub.get("description"),
                        createdDateTime=sub.get("createdDateTime"),
                        lastModifiedDateTime=sub.get("lastModifiedDateTime"),
                    ),
                    message=f"Subsite found for company '{request.company_url_slug}'",
                )
        return SubsiteLookupResponse(
            exists=False,
            subsite=None,
            message=f"No subsite found for company '{request.company_url_slug}'. Create new subsite.",
        )
    except HTTPException:
        raise
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Subsite lookup failed: {exc}") from exc

# ------------------------------------------------------------------
# Error handlers
# ------------------------------------------------------------------

@app.exception_handler(httpx.HTTPStatusError)
async def httpx_exception_handler(request, exc):
    return JSONResponse(
        status_code=exc.response.status_code,
        content={"detail": f"External API error: {exc.response.text}", "error_code": "ExternalAPIError"},
    )

# ------------------------------------------------------------------
# Main entry point
# ------------------------------------------------------------------

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
