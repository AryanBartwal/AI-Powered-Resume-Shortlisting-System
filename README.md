# 🚀 AI-Powered Resume Shortlisting System

> An intelligent MERN Stack application with Machine Learning for automated resume screening and candidate ranking

## 📋 Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Tech Stack](#tech-stack)
- [System Architecture](#system-architecture)
- [Quick Start](#quick-start)
- [Usage Guide](#usage-guide)
- [API Documentation](#api-documentation)
- [ML Algorithm](#ml-algorithm)
- [Project Structure](#project-structure)

---

## 🎯 Overview

This system automates the resume screening process using AI/ML, helping recruiters save time and reduce bias. It analyzes resumes against job descriptions using NLP and ranks candidates based on skill match, experience, education, and semantic similarity.

**Key Highlights:**
- 🤖 BERT-based semantic similarity for context understanding
- 🎨 Modern React UI with intuitive dashboard
- 📊 Real-time candidate ranking with detailed insights
- 📁 Supports PDF and DOCX resume formats
- 🔒 In-memory storage (no database setup needed)
- 🚀 One-click deployment with Docker support

---

## ✨ Features

### For Recruiters (Admin)
- ✅ Post and manage job openings
- ✅ Set required skills for each position
- ✅ Generate AI-powered candidate rankings
- ✅ View detailed match analysis (skills, projects, experience)
- ✅ Filter dashboard to show only own job postings

### For Job Seekers
- ✅ Upload resume and apply to specific jobs
- ✅ Automatic resume parsing (PDF/DOCX)
- ✅ Match score with detailed skill breakdown
- ✅ Apply to multiple positions

### ML-Powered Analysis
- ✅ **Skill Matching (65%)** - Detects skill variations (React/React.js, DSA/Data Structures)
- ✅ **BERT Semantic Similarity (12%)** - Advanced NLP for context understanding
- ✅ **Experience/Education Match (10%)** - Dynamic requirement matching
- ✅ **Project Relevance (8%)** - Analyzes relevant work/projects
- ✅ **Role Keywords (3%)** - Detects seniority levels
- ✅ **Certifications (2%)** - Bonus for relevant certifications

---

## 🛠️ Tech Stack

### Frontend
- **React 18** - Modern UI library
- **React Router** - Client-side routing
- **Axios** - HTTP client
- **CSS3** - Responsive design with gradient themes

### Backend
- **Node.js** - JavaScript runtime
- **Express.js** - Web framework
- **Multer** - File upload handling
- **JWT** - Authentication
- **MongoDB** (optional) - Data persistence

### ML Pipeline
- **Python 3.x** - Programming language
- **Flask** - Lightweight web framework
- **scikit-learn** - TF-IDF vectorization
- **Sentence-BERT** - Semantic similarity
- **PyPDF2** - PDF parsing
- **python-docx** - DOCX parsing

### DevOps
- **Docker** - Containerization
- **Docker Compose** - Multi-container orchestration

---

## 🏗️ System Architecture

```
┌─────────────────┐      ┌─────────────────┐      ┌─────────────────┐
│   Frontend      │      │    Backend      │      │   ML Pipeline   │
│   React:3000    │◄────►│  Express:5000   │◄────►│   Flask:5001    │
│                 │      │                 │      │                 │
│ - Job Posting   │      │ - Job APIs      │      │ - Resume Parse  │
│ - Resume Upload │      │ - Resume APIs   │      │ - TF-IDF        │
│ - Rankings      │      │ - Ranking APIs  │      │ - BERT Embed    │
└─────────────────┘      └─────────────────┘      └─────────────────┘
```

**Data Flow:**
1. Recruiter posts job with required skills
2. Candidates upload resumes
3. ML pipeline parses resumes (PDF/DOCX → text)
4. BERT + TF-IDF calculate semantic similarity
5. Skill matching detects required skills in resume
6. Combined scoring: 65% skills + 12% semantic + 10% experience + 8% projects + 5% other
7. Ranked results displayed in dashboard

---

## 🚀 Quick Start

### Prerequisites
- Node.js (v14+)
- Python (v3.8+)
- npm or yarn

### Installation

**1. Clone & Install:**
```powershell
cd "MAJOR Project 2"

# Install backend dependencies
cd backend
npm install

# Install frontend dependencies
cd ../frontend
npm install

# Install ML dependencies
cd ../ml-pipeline
pip install -r requirements.txt
```

**2. Start All Services:**
```powershell
# One command startup (from project root)
.\START.ps1
```

Or start manually:
```powershell
# Terminal 1 - Backend
cd backend
npm run dev

# Terminal 2 - ML Pipeline
cd ml-pipeline
python src/app.py

# Terminal 3 - Frontend
cd frontend
npm start
```

**3. Access Application:**
- Frontend: http://localhost:3000
- Backend API: http://localhost:5000
- ML Pipeline: http://localhost:5001

---

## 📖 Usage Guide

### Step 1: Post a Job
1. Navigate to http://localhost:3000
2. Click **"Post Job"** tab
3. Fill in details:
   - Job Title: `Software Engineer (Full Stack)`
   - Company: `Tech Corp`
   - Description: Detailed requirements
   - Skills: `Java, Python, React, Node.js, SQL`
4. Click **"Post Job"**

### Step 2: Upload Resumes
1. Click **"Upload Resume"** tab or **"Apply Now"** on job card
2. Enter candidate details
3. Select job from dropdown
4. Upload PDF/DOCX resume
5. Submit

### Step 3: View Rankings
1. Click **"Dashboard"** tab
2. Select a job
3. Click **"Generate Rankings"**
4. View ranked candidates with:
   - Match percentage (e.g., 68%)
   - Matched skills count
   - Missing skills
   - Rank position

---

## 🔧 API Documentation

### Jobs API
- `POST /api/jobs` - Create new job posting
- `GET /api/jobs` - List all jobs
- `GET /api/jobs/:id` - Get single job details
- `PUT /api/jobs/:id` - Update job
- `DELETE /api/jobs/:id` - Delete job

### Resumes API
- `POST /api/resumes/upload` - Upload resume (multipart/form-data)
- `POST /api/resumes/match-jobs` - Match resume against all jobs
- `GET /api/resumes` - List all resumes
- `GET /api/resumes/job/:jobId` - Get resumes for specific job

### Rankings API
- `POST /api/rankings/generate/:jobId` - Generate AI rankings for job
- `GET /api/rankings/job/:jobId` - Get rankings for job

### ML Pipeline API
- `POST /parse/resume` - Parse PDF/DOCX resume
- `POST /ranking/calculate` - Calculate similarity scores

---

## 🧠 ML Algorithm

### Scoring Components (100% Total)

**1. Skills Match (65%)**
- Detects skill variations (React = React.js, DSA = Data Structures)
- Uses word boundary matching
- N-gram support for multi-word skills

**2. BERT Semantic Similarity (12%)**
- Uses Sentence-BERT (`all-MiniLM-L6-v2`)
- Understands context and meaning
- Falls back to TF-IDF if BERT unavailable

**3. Experience & Education (10%)**
- Detects years of experience (0-2, 2-5, 5+)
- Matches education level (B.Tech, M.Tech, PhD)
- Dynamic requirement matching

**4. Project Relevance (8%)**
- Analyzes project keywords
- Compares with job requirements
- Generic for all fields

**5. Role Keywords (3%)**
- Seniority levels (Fresher, Junior, Senior)
- Role types (Manager, Lead, Director)

**6. Certifications (2%)**
- AWS, Azure, MongoDB, etc.
- Bonus points for relevant certs

### Formula
```
Score = (Skills × 0.65) + (BERT × 0.12) + (Experience × 0.10) + 
        (Projects × 0.08) + (Keywords × 0.03) + (Certs × 0.02)
```

### Interpretation
- **60-75%** = Strong match, definitely interview
- **50-60%** = Good match, consider
- **40-50%** = Decent match, review
- **<40%** = Weak fit

---

## 📁 Project Structure

```
MAJOR Project 2/
├── backend/               # Node.js + Express API
│   ├── src/
│   │   ├── server.js     # Main server
│   │   ├── config/       # Database config
│   │   ├── models/       # Mongoose models
│   │   ├── routes/       # API routes
│   │   ├── middleware/   # Auth middleware
│   │   └── storage/      # In-memory storage
│   ├── uploads/          # Uploaded resumes
│   └── package.json
│
├── frontend/             # React application
│   ├── src/
│   │   ├── App.js       # Main app component
│   │   ├── pages/       # Page components
│   │   ├── components/  # Reusable components
│   │   ├── contexts/    # React contexts
│   │   └── styles/      # CSS files
│   ├── build/           # Production build
│   └── package.json
│
├── ml-pipeline/         # Python + Flask ML service
│   ├── src/
│   │   ├── app.py      # Flask server
│   │   ├── routes/     # API routes
│   │   └── utils/      # ML utilities
│   │       ├── parser.py      # Resume parsing
│   │       └── similarity.py  # ML algorithms
│   └── requirements.txt
│
├── docker/              # Docker configuration
│   ├── docker-compose.yml
│   ├── Dockerfile.backend
│   ├── Dockerfile.frontend
│   └── Dockerfile.ml
│
├── START.ps1           # Launch script
└── README.md           # Documentation
```

---

## 🐳 Docker Deployment (Optional)

```bash
# Build and run all services
docker-compose up --build

# Stop all services
docker-compose down
```

Services will be available at:
- Frontend: http://localhost:3000
- Backend: http://localhost:5000
- ML Pipeline: http://localhost:5001

---

## ⚡ Troubleshooting

### Port Already in Use
```powershell
# Kill process on port 5000
Get-Process -Id (Get-NetTCPConnection -LocalPort 5000).OwningProcess | Stop-Process -Force

# Kill process on port 3000
Get-Process -Id (Get-NetTCPConnection -LocalPort 3000).OwningProcess | Stop-Process -Force
```

### ML Service Issues
```powershell
cd ml-pipeline
pip install --upgrade pip
pip install -r requirements.txt
python src/app.py
```

### Frontend Build Issues
```powershell
cd frontend
rm -rf node_modules
npm install
npm start
```

---

## 🎓 Learning Outcomes

This project demonstrates:
- ✅ **MERN Stack Development** - Full-stack JavaScript
- ✅ **Machine Learning Integration** - NLP + ML pipelines
- ✅ **RESTful API Design** - Best practices
- ✅ **File Handling** - PDF/DOCX parsing
- ✅ **Real-time Data Processing** - Efficient algorithms
- ✅ **Modern UI/UX** - React best practices
- ✅ **Microservices Architecture** - Service separation
- ✅ **Production-Ready Code** - Error handling, validation

---

## 🚀 Future Enhancements

- 🔐 Add user authentication with roles
- 📊 Advanced analytics dashboard
- 📧 Email notifications for candidates
- 🌐 Deploy to cloud (AWS/Heroku/Azure)
- 📱 Mobile-responsive improvements
- 🤖 LinkedIn profile integration
- 🧪 Unit and integration tests
- 📈 Bias detection in hiring

---

## 📄 License

MIT License - Feel free to use this project for learning and development.

---

## 👨‍💻 Author

Built with ❤️ for automating resume screening and making hiring fair and efficient!

---

## 🙏 Acknowledgments

- **scikit-learn** - Machine learning library
- **Sentence-BERT** - Semantic similarity
- **React Team** - Amazing frontend framework
- **Express.js** - Fast backend framework
- **Flask** - Lightweight Python framework

---

**⭐ Star this project if you found it useful!**
- Stored in `/uploads` directory

### ✅ Real-Time Rankings
- Instant similarity calculation
- Skill matching visualization
- Top candidates displayed first
- Match percentage breakdown

### ✅ Beautiful UI
- Modern gradient design (purple theme)
- Responsive layout
- Tabbed navigation
- Success/error notifications

---

## 🚀 PRODUCTION DEPLOYMENT

To deploy to production:

1. **Add MongoDB** (replace in-memory storage)
2. **Add Authentication** (JWT tokens)
3. **Deploy Backend** (Heroku, AWS, Azure)
4. **Deploy Frontend** (Vercel, Netlify)
5. **Deploy ML Pipeline** (Docker container)
6. **Add File Storage** (AWS S3, Azure Blob)

---

## 📊 SAMPLE DATA

### Sample Job Post
```json
{
  "title": "Senior Full Stack Developer",
  "company": "Tech Corp",
  "description": "We are looking for an experienced full stack developer...",
  "requiredSkills": ["React", "Node.js", "MongoDB", "AWS", "Docker"]
}
```

### Sample Resume Upload
- Name: John Doe
- Email: john@example.com
- File: john_resume.pdf
- Job: Senior Full Stack Developer

---

## 🎓 TECHNOLOGIES USED

- **Frontend**: React 18, React Router, Axios
- **Backend**: Node.js, Express, Multer
- **ML**: Python, Flask, scikit-learn, PyPDF2
- **Algorithms**: TF-IDF, Cosine Similarity
- **Storage**: In-Memory (Map-based)

---

## 📝 LICENSE

Open source for educational purposes.

---

## 👨‍💻 DEVELOPMENT NOTES

- All API responses follow `{success: boolean, data: any}` format
- Resume IDs are auto-incremented integers
- Rankings automatically expire when job is updated
- Frontend uses React Hooks (useState, useEffect)
- ML pipeline returns top 10 candidates by default

---

## 🌟 NEXT STEPS

1. **Test the workflow**: Post job → Upload resumes → Generate rankings
2. **Customize styling**: Edit `/frontend/src/styles/Home.css`
3. **Improve ML model**: Add more sophisticated NLP (BERT, transformers)
4. **Add features**: Email notifications, candidate profiles, interview scheduling
5. **Deploy to cloud**: Make it accessible to real users!

---

**🎉 YOUR AI-POWERED RESUME SHORTLISTING SYSTEM IS COMPLETE!**

Open http://localhost:3000 and start screening resumes with AI! 🚀
