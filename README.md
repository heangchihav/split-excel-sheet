# Split Excel Sheet Web Tool

An enterprise-grade, batch Excel sheet splitting web tool built with a modern **Next.js 16** frontend and a dedicated **Spring Boot 3.3 (Java 21 + Apache POI)** backend.

---

## Architecture Overview

```
split-excel-sheet/
├── backend/                   # Spring Boot 3.3 + Apache POI Service
│   ├── src/main/java/         # Controllers, Services, DTOs
│   ├── src/main/resources/    # application.yml
│   ├── pom.xml                # Apache POI 5.3.0, Java 21, Spring Web
│   ├── Dockerfile             # Multi-stage Eclipse Temurin 21 + Khmer fonts
│   └── mvnw                   # Maven Wrapper (no local Java install required)
├── components/                # Next.js React UI components
├── lib/                       # Utility functions & API client with client fallback
├── app/                       # Next.js App Router & Layout
├── docker-compose.yml         # Unified orchestration for frontend & backend
└── Dockerfile                 # Frontend Next.js production container
```

---

## Key Features

1. **Enterprise Apache POI Engine**:
   - Native support for both legacy `.xls` (BIFF8) and modern `.xlsx` (OpenXML).
   - 100% preservation of cell styles, Khmer Unicode fonts (`Khmer OS Battambang`, `Khmer OS Muol`, etc.), borders, fills, formulas, and merged regions.
   - Zero Microsoft Excel repair or recovery warnings.
2. **Batch & Multi-File Processing**:
   - Upload multiple workbooks at once.
   - Configure custom folder names, prefixes, separators, sequence numbering, or original sheet names.
   - Download as individual single-sheet workbooks or as organized ZIP archives.
3. **Decoupled Architecture with High Resilience**:
   - Spring Boot handles heavy document manipulation on port `8080`.
   - Next.js frontend runs on port `3000` with transparent API rewrites.
   - Dual-mode client fallback ensures the frontend remains functional even when running offline.

---

## Quick Start (Docker Compose)

The easiest way to run the complete stack is with Docker (no Java or Node installation required on your host system):

```bash
docker compose up --build
```

- **Frontend UI**: [http://localhost:3000](http://localhost:3000)
- **Spring Boot Backend**: [http://localhost:8080](http://localhost:8080)
- **Health Check**: [http://localhost:8080/api/health](http://localhost:8080/api/health)

---

## Local Development (Without Docker)

### 1. Spring Boot Backend
```bash
cd backend
./mvnw spring-boot:run
```
*(Requires Java 21 on your host system)*

### 2. Next.js Frontend
```bash
# In the project root:
npm install
npm run dev
```
Open [http://localhost:3000](http://localhost:3000) in your browser.

---

## REST API Endpoints

- `GET /api/health`: Health status and engine metadata.
- `POST /api/excel/inspect`: Inspects an uploaded workbook and returns sheet names, dimensions, and preview rows.
- `POST /api/excel/split-single`: Isolates a single worksheet into an independent `.xlsx` file.
- `POST /api/excel/split-zip`: Splits selected sheets from a workbook into a folder inside a ZIP archive.
- `POST /api/excel/split-batch-zip`: Splits multiple master workbooks and bundles them into organized master folders within a master ZIP archive.
