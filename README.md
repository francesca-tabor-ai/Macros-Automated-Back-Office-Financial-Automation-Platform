# GenAI Data Cleaning

AI-powered data transformation platform that separates rule generation from execution. Claude defines the cleaning rules, your code enforces them deterministically.

---

## Overview

This platform provides a robust, auditable approach to data cleaning where:
- **AI defines the rules** through natural language intent
- **Code enforces them** with deterministic execution
- **Data becomes reliable** with full audit trails and explainability

Perfect for enterprise analytics, data governance, and regulated environments where transparency and reproducibility matter.

---

## Tech Stack

### Backend
- Python 3.10+
- FastAPI
- Pandas
- Pydantic
- Pytest

### Frontend
- Next.js (App Router)
- React
- TypeScript

### AI Workflow
- **Claude** → rule generation  
- **Cursor** → implementation, iteration, and refactoring

---

## Getting Started

### Prerequisites
- Python 3.10+
- Node.js 18+
- npm or pnpm
- (Optional) Claude API key

### Clone the Repository
```bash
git clone https://github.com/your-org/genai-data-cleaning.git
cd genai-data-cleaning
```

### Backend Setup
```bash
cd backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Run the API:
```bash
uvicorn app.main:app --reload
```

API available at: `http://localhost:8000`

### Frontend Setup
```bash
cd frontend
npm install
npm run dev
```

Frontend available at: `http://localhost:3000`

---

## API Endpoints

### Upload Dataset
```http
POST /datasets/upload
```

### Preview Raw Data
```http
GET /datasets/{dataset_id}/preview
```

### Apply Transformation Spec
```http
POST /datasets/{dataset_id}/transform
```

### Job Status & Logs
```http
GET /jobs/{job_id}
```

### Export Clean Data
```http
GET /datasets/{dataset_id}/export?format=csv|xlsx
```

---

## Transformation Spec Example
```json
{
  "version": "1.0",
  "dataset_name": "crm_contacts",
  "rules": [
    {
      "id": "R1",
      "type": "standardize",
      "field": "email",
      "params": {
        "lowercase": true,
        "trim": true
      }
    },
    {
      "id": "R2",
      "type": "deduplicate",
      "params": {
        "keys": ["email"],
        "tie_breaker": "most_recent"
      }
    }
  ]
}
```

---

## Explainability & Auditing

Every transformation produces:
- Rule-level metrics
- Sample before/after values
- Dropped and quarantined row counts
- Deterministic execution logs

Designed for:
- Enterprise analytics
- Data governance
- Regulated environments

---

## Development Workflow (Cursor + Claude)

1. Upload or preview a dataset
2. Describe cleaning intent in natural language
3. Claude generates a Transformation Spec
4. Apply the spec via the API
5. Inspect diffs and audit logs
6. Export the clean dataset

**Rules evolve. Code stays stable.**

---

## Testing

Run backend tests:
```bash
pytest
```

Test coverage includes:
- Rule correctness
- Deduplication edge cases
- Schema enforcement
- Quarantine logic

---

## Security & Data Handling

- Data remains in your environment
- AI operates on samples or metadata only (configurable)
- No silent or irreversible mutations
- Full audit trail for every transformation

---

## Roadmap

- [ ] Native Claude API integration
- [ ] Prebuilt domain rule packs (CRM, Billing, Network)
- [ ] Data quality scoring
- [ ] BI and warehouse connectors
- [ ] Streaming ingestion support

---

## Contributing

Contributions are welcome:
- New rule types
- Domain-specific transformations
- Performance optimizations
- UX improvements

Please open an issue before submitting major changes.

---

## License

MIT License

See [LICENSE](LICENSE) for details.

---

## TL;DR

**AI defines the rules.**  
**Code enforces them.**  
**Data becomes reliable.**
