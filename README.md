# 🎯 SlideGen AI – AI-Powered PPT Generator

## 🚀 Overview

SlideGen AI is a full-stack web application that generates professional PowerPoint presentations using Artificial Intelligence. It allows users to create, edit, preview, and download PPT files instantly using AI-generated structured content.

---

## 🧰 Tech Stack

### Frontend

* React 18
* Vite
* Tailwind CSS
* Zustand (State Management)
* Framer Motion (Animations)

### Backend

* FastAPI (Python)
* Groq API (AI Content Generation)
* python-pptx (PPT Creation)

---

## 📁 Project Structure

```
slidegen-ai/
│
├── backend/
│   ├── main.py                 # FastAPI entry point
│   ├── routes/                # API routes (generate, edit, export)
│   ├── services/              # Business logic & AI integration
│   ├── models/                # Request/response schemas
│   ├── utils/                 # Helper functions
│   └── output/                # Generated PPT files
│
├── frontend/
│   ├── src/
│   │   ├── components/        # UI components
│   │   ├── pages/             # App pages (generator, onboarding)
│   │   ├── store/             # Zustand state management
│   │   ├── utils/             # Helper functions
│   │   └── App.jsx            # Main app component
│   ├── public/
│   └── index.html
│
├── .env                       # Environment variables
├── requirements.txt           # Python dependencies
├── package.json               # Frontend dependencies
└── README.md
```

---

## ⚙️ How to Run the Project

### 🔧 Backend Setup

1. Clone the repository:

```bash
git clone https://github.com/your-username/slidegen-ai.git
cd slidegen-ai
```

2. Create virtual environment:

```bash
python -m venv venv
source venv/bin/activate   # Linux/macOS
venv\Scripts\activate    # Windows
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Set environment variables:

* Create `.env` file
* Add:

```
GROQ_API_KEY=your_api_key_here
```

5. Run backend server:

```bash
uvicorn backend.main:app --reload
```

Backend runs at:
👉 [http://localhost:8000](http://localhost:8000)

---

### 💻 Frontend Setup

1. Navigate to frontend:

```bash
cd frontend
```

2. Install dependencies:

```bash
npm install
```

3. Run frontend:

```bash
npm run dev
```

Frontend runs at:
👉 [http://localhost:5173](http://localhost:5173)

---

## 🔌 API Endpoints

| Endpoint         | Method | Description                  |
| ---------------- | ------ | ---------------------------- |
| /generate        | POST   | Generate full presentation   |
| /slides/edit     | POST   | Edit a single slide using AI |
| /export          | POST   | Export PPT after edits       |
| /download/{file} | GET    | Download PPT file            |
| /health          | GET    | Check API status             |
| /history         | GET    | Get recent presentations     |

---

## ✨ Features

* AI-powered PPT generation
* Slide editing (manual + AI)
* Real-time preview
* PPT download support
* Theme customization
* Local history storage

---

## 🧠 AI Workflow

1. User enters topic
2. Backend sends prompt to Groq API
3. AI returns structured JSON
4. Backend normalizes slides
5. python-pptx generates PPT file
6. Frontend displays and allows edits

---

## 🧪 Default Settings

* Slides: 7 (range: 5–12)
* Tone: Professional
* Audience: Business
* Theme: Modern

---

## 🚨 Notes

* No authentication system (local storage used)
* Requires stable internet for AI generation
* Generated PPT files stored on backend

---

## 👨‍💻 Author

Pratham Mishra

---

## ⭐ Support

If you like this project, give it a ⭐ on GitHub!
