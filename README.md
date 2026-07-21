# CareerForge Pro

AI-powered career tools with cryptocurrency payments. Generate professional resumes, optimize for ATS, prepare for interviews, and negotiate salaries.

## Features

### Free Tools
- **Free ATS Resume Score** - Upload your resume or paste email to get instant ATS scoring and keyword gap analysis

### Paid Tools ($19-$49)
- **AI Resume Generator** ($49) - AI-written resume tailored to your role + cover letter + PDF download
- **Resume Optimizer** ($49) - ATS scoring, keyword analysis, rewritten summary
- **Interview Prep** ($29) - Role-specific questions with STAR method examples
- **Salary Negotiator** ($19) - Market rate benchmarking + negotiation scripts

## Tech Stack

- **Backend**: Flask (Python)
- **Server**: Gunicorn WSGI server
- **Database**: SQLite (for orders and email leads)
- **AI**: Anthropic Claude 3.5 Sonnet (for resume generation)
- **Payments**: NOWPayments (Bitcoin, Ethereum, USDT, Solana)
- **PDF**: ReportLab

## Setup

```bash
# Install dependencies
pip install -r requirements.txt

# Set environment variables (copy and edit)
cp .env.example .env

# Run development server
python app.py

# Or run with gunicorn (production)
gunicorn -c gunicorn.conf.py app:app
```

## Production Deployment

### Linux/macOS
```bash
# Install dependencies
pip install -r requirements.txt

# Set environment variables
export ANTHROPIC_API_KEY=your-key-here
export SECRET_KEY=$(python -c "import secrets; print(secrets.token_hex(32))")
export FLASK_ENV=production

# Start server
gunicorn -c gunicorn.conf.py app:app
```

### Windows
```cmd
# Set environment variables
set ANTHROPIC_API_KEY=your-key-here
set SECRET_KEY=your-secret-here
set FLASK_ENV=production

# Start server
gunicorn -c gunicorn.conf.py app:app
```

## Environment Variables

| Variable | Required | Description |
|----------|----------|-------------|
| `SECRET_KEY` | Production | Flask session secret key |
| `ANTHROPIC_API_KEY` | AI tools | For AI resume generation |
| `ADMIN_SECRET` | Optional | For accessing `/api/emails` endpoint |
| `NOWPAYMENTS_API_KEY` | Optional | For real crypto payment verification |
| `FLASK_ENV` | Production | Set to "production" to disable debug |

## API Endpoints

| Endpoint | Method | Description |
|----------|--------|-------------|
| `/api/services` | GET | Get available services |
| `/api/email-capture` | POST | Capture email + resume for free ATS score |
| `/api/upload-resume` | POST | Upload PDF/DOCX resume file |
| `/api/orders` | POST | Create payment order |
| `/api/generate` | POST | AI resume generation |
| `/api/generate-result` | POST | Generate tool results |
| `/api/pdf` | POST | Create downloadable PDF |
| `/api/download/<filename>` | GET | Download PDF |
| `/api/emails` | GET | List captured emails (requires admin secret) |
| `/health` | GET | Health check endpoint |

## SEO Pages

- `/ats-resume-checker` - Free ATS checker landing
- `/ai-cover-letter-generator` - AI cover letter page
- `/interview-questions-generator` - Interview prep page
- `/resume-score` - Resume scoring page
- `/resume-keywords` - Keyword analyzer page

## Security Features

- Rate limiting (200/day, 50/hour default)
- Security headers (X-Frame-Options, CSP, HSTS, etc.)
- Input validation on all endpoints
- File size limits (8MB max)
- Secure filename handling

## Notes

- AI resume generation requires `ANTHROPIC_API_KEY` environment variable
- Without API key, only free ATS scoring and interview/salary tools work
- NOWPayments button IDs are configured in the app for crypto payments
- All resume data is processed, no persistent storage of user resumes
- Use gunicorn for production - Flask dev server is not suitable for production