# Gent-Chat Deployment Guide

## Prerequisites
- Docker & Docker Compose
- Backend CEM running on `backend_cem_default` network
- Environment variables configured

## Quick Start

### 1. Setup Environment
```bash
cp .env.example .env.local
# Edit .env.local with your credentials
```

### 2. Build & Run
```bash
# Build image
docker-compose build

# Start chatbot
docker-compose up -d

# Check logs
docker-compose logs -f chatbot
```

### 3. Stop
```bash
docker-compose down
```

## Environment Variables

| Variable | Description | Example |
|----------|-------------|---------|
| `CEM_API_URL` | CEM Backend API URL | `http://backend-cem:3001/api` |
| `CEM_API_TOKEN` | JWT Token for CEM API | `eyJhbGc...` |
| `GEMINI_API_KEY` | Google Gemini API Key | `AIza...` |
| `AZURE_CLIENT_ID` | Azure App Client ID | `997e1b06...` |
| `AZURE_CLIENT_SECRET` | Azure App Secret | `N4E8Q~...` |
| `AZURE_TENANT_ID` | Azure Tenant ID | `c5fc1b2a...` |

## Features

### CEM Integration
Chatbot can answer questions about:
- พนักงาน (Employees)
- โครงการ (Projects/Tasks)
- การลา (Leave Requests)
- การจองรถ (Car Bookings)
- งานประจำวัน (Daily Work)

### Commands
- `clear` - Clear conversation history
- `model <name>` - Switch AI model
- `/broadcast <message>` - Broadcast message

## API Endpoints

Chatbot connects to CEM API:
- `GET /api/users` - Get users
- `GET /api/tasks` - Get tasks
- `GET /api/leave` - Get leave requests
- `GET /api/car-booking` - Get car bookings
- `GET /api/daily-work` - Get daily work

## Troubleshooting

### Cannot connect to CEM API
```bash
# Check if backend is running
docker ps | grep backend-cem

# Check network
docker network ls | grep backend_cem_default

# Test API connection
docker exec gent-chatbot wget -O- http://backend-cem:3001/api/test-db
```

### Chatbot not responding
```bash
# Check logs
docker-compose logs chatbot

# Restart
docker-compose restart chatbot
```

## Architecture

```
User (Teams) → Webhook → Gemini AI → CEM API → Database
                  ↓
            Response with CEM Data
```

## Port
- Chatbot: `3002` (external) → `3000` (internal)
- Backend CEM: `3001`
