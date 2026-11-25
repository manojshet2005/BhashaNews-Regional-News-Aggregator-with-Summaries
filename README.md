# BhashaNews — Regional News Aggregator with Summaries

Case P57: Media project for roll number 57.

## Stack
- React frontend
- Python FastAPI backend
- PostgreSQL through SQLAlchemy; SQLite fallback for local development
- Transformers-ready summarisation interface with a clearly identified local demonstration fallback
- Docker and Docker Compose
- pytest-cov, Selenium, GitHub Actions, Jenkins

## Features
Readers can browse, search, filter, and open original article URLs. Editors can switch to an editor view and trigger summaries. The API supports article CRUD, summaries, health, metrics, and language/category options.

## Start locally
```powershell
cd backend
python -m pip install -r requirements.txt
python seed.py
uvicorn app:app --host 127.0.0.1 --port 8000
```

In a second terminal:
```powershell
cd frontend
npm install
npm run dev -- --host 127.0.0.1
```

Open `http://127.0.0.1:5173`. API documentation is at `http://127.0.0.1:8000/docs`.

## PostgreSQL
Copy `.env.example` to `.env` and set `DATABASE_URL`. Start the full stack with:
```powershell
docker compose up --build
```

## Tests
```powershell
cd backend
pytest
```

Measured local suite: 11 tests passed and 94% total statement coverage.

## AI limitation
The implementation includes a documented extractive demonstration summariser. A pretrained Transformers model can be integrated behind the existing summary function when model weights and runtime dependencies are available. The local fallback is not represented as an actual pretrained Transformer execution.
