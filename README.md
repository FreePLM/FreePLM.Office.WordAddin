# FreePLM Word Add-in

## Overview

This VSTO Word Add-in provides PLM (Product Lifecycle Management) functionality directly within Microsoft Word. It communicates with the FreePLM service via REST API.

## Architecture

```
Word Add-in (.NET Framework 4.8)
    ↓ HTTP/REST
FreePLM Service (localhost:5000)
    ↓
Backend (.NET 10)
    ↓
SQLite Database + Local File Storage
```

## Features Implemented

### 1. Custom Ribbon UI
- **FreePLM Tab** in Word ribbon
- **Document Control Group**:
  - Check Out button
  - Check In button
  - Cancel Check Out button
- **Document Information Group**:
  - View Properties button
  - Refresh button

### 2. API Client (`PLMApiClient.cs`)
- Communicates with FreePLM service on `localhost:5000`
- Methods:
  - `GetDocumentAsync()` - Get document metadata
  - `DownloadFileAsync()` - Download file content
  - `CheckOutAsync()` - Check out document
  - `CheckInAsync()` - Check in document with new revision
  - `CancelCheckOutAsync()` - Cancel checkout
  - `ChangeStatusAsync()` - Change document status
  - `IsServiceAvailableAsync()` - Health check

### 3. Data Models (DTOs)
- `DocumentStatus` enum - Private, InWork, Frozen, Released, Obsolete
- `DocumentReadDto` - Document metadata
- `CheckOutDto` / `CheckOutResultDto` - Check out operation
- `CheckInDto` / `CheckInResultDto` - Check in operation
- `ApiErrorResponse` - Error handling

### 4. Exception Handling
- `PLMApiException` - Custom exception for API errors
- Helper properties:
  - `IsDocumentLocked`
  - `IsLockNotHeld`
  - `IsNotFound`

## User Workflow

### Check Out a Document
1. User clicks **Check Out** button
2. Prompted for ObjectId (or reads from document properties)
3. Prompted for comment
4. API call to check out document
5. File downloaded to temp location
6. Opened in Word
7. ObjectId stored in document custom properties

### Check In a Document
1. User makes changes to document
2. User clicks **Check In** button
3. Document is saved
4. Prompted for check-in comment
5. Asked if Major or Minor revision
6. File uploaded to PLM system
7. New revision created
8. Document closed

### Cancel Check Out
1. User clicks **Cancel** button
2. Confirms cancellation
3. Check out cancelled in PLM system
4. Document closed without saving

### View Properties
1. User clicks **Properties** button
2. Dialog shows:
   - ObjectId, FileName, Revision
   - Status, Owner, Group, Role, Project
   - Created/Modified dates and users
   - Check out status

## Custom Document Properties

The add-in stores PLM metadata in Word's custom document properties:
- `PLM_ObjectId` - Document's PLM ObjectId
- `PLM_Revision` - Current revision
- `PLM_CheckedOut` - Check out status (true/false)

## Dependencies

### NuGet Packages
- `Newtonsoft.Json` 13.0.3 - JSON serialization
- `System.Net.Http` (built-in) - HTTP client

### Office References
- Microsoft.Office.Interop.Word
- Microsoft.Office.Core
- Microsoft.Office.Tools.Word
- VSTO Runtime

## Configuration

### Service URL
Default: `http://localhost:5000`

Can be changed in `PLMApiClient` constructor.

## Error Handling

### Service Not Available
If the FreePLM service is not running, user sees:
> "Cannot connect to FreePLM service. Please ensure the FreePLM service is running on localhost:5000"

### Document Locked
If document is already checked out by another user:
> "Document is already checked out. [locked by user]"

### Not Checked Out
If trying to check in a document that's not checked out:
> "This document is not checked out from FreePLM."

## Next Steps - Backend Service

The backend service needs to implement these endpoints:

### Document Endpoints
- `GET /api/documents/{objectId}` - Get document info
- `GET /api/documents/{objectId}/content` - Download file
- `POST /api/documents` - Create new document

### CheckOut/CheckIn Endpoints
- `POST /api/checkout` - Check out document
- `POST /api/checkin` - Check in document (multipart form with file)
- `POST /api/checkout/cancel` - Cancel checkout
- `GET /api/checkout/{objectId}/status` - Get lock status

### Workflow Endpoints
- `POST /api/workflow/status` - Change status

### Health Check
- `GET /api/health` - Service health check

## Status

✅ Word Add-in - **COMPLETE**
⏳ Backend Service - **NEXT STEP**
⏳ Database Schema - **NEXT STEP**
⏳ Storage Provider - **NEXT STEP**
