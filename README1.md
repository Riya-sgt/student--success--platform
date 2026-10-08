# 🎓 EduNexus AI - Smart Campus Analytics & Student Success Platform

Built for the **Hacxlerate - Bytexl Hackathon** (KPMG in India problem statement: *AI-Powered Student Analytics and Success Platform*).

**🔗 Live Demo:** [Open EduNexus AI on Streamlit Cloud](https://student--success--platform-ajpnzmgfwmmgzdhcld7zq5.streamlit.app/)

---

## 📌 Problem Statement

**Smart Campus Analytics: Predict, Optimize & Improve Student Success**

Colleges generate large amounts of data through attendance, internal assessments, exams, LMS activity, placement participation, extracurricular activities and student feedback. But this data lives in disconnected systems and is rarely used proactively. Because of this:

- Faculty find out about a struggling student only when it is already too late.
- Students with low attendance or low engagement are not noticed in time.
- Placement officers cannot easily see which students lack the skills for their target job.
- Administrators have dashboards, but no clear "what should we do next?" guidance.

The challenge: build an Analytics & AI-powered platform that identifies students who may need intervention and helps administrators take data-driven action, going beyond a simple dashboard.

## 💡 Our Solution

**EduNexus AI** brings all student data into one unified view. It gives every student a **Success Score** and a **risk level**, explains **why** a student is at risk, groups students into **segments** for targeted action, and helps faculty act quickly through an intervention workflow.

---

## ✨ Features

| Tab | What it does |
|---|---|
| 🏠 **Command Center** | Campus overview: total students, average Success Score, High-risk and Critical counts, risk distribution and department-wise risk. Also shows the **AI Cohort Matrix**, which groups students into 6 segments: Low Attendance, High Academic / Low Placement, Consistent All-Rounders, Low Engagement, Academic Strugglers and High Performers. |
| 🎓 **Student 360** | Complete profile of one student, a "Why at risk?" explanation, and an inbox for messages from teachers. |
| 💼 **Skill Gap** | Compares the student's current skills with the skills required for the target role (shows missing skills) and a skill radar against the cohort baseline. |
| 🧠 **Mental Health & Wellness** | Stress and wellness indicators with suggested next steps. Hidden from Placement Officers. |
| 🤖 **AI Copilot** | AI-generated outreach email that faculty can edit and send, plus a live Intervention Tracker (Pending / In Progress / Completed). |
| 🎛 **What-If Simulator** | Shows how the Success Score changes if attendance reaches 80% or skills improve by 10 points. |

### 🔐 Role-Based Access

- **Administrator** - full access to all students
- **Faculty / Mentor** - sees only students of the selected department
- **Placement Officer** - sees only High/Critical placement-risk students (wellness data is hidden)
- **Student** - signs in with Student ID and PIN and sees only their own data

### 📂 Dataset

- The full **1000-student dataset (45 columns)** is built into the app, so it works right after deployment.
- You can also upload your own `.xlsx` file from the sidebar to replace the data.

---

## 🛠 Tech Stack

- **Streamlit** - web app and UI
- **Pandas & NumPy** - data processing
- **Plotly** - interactive charts
- **OpenPyXL** - reading Excel files

---

## 🚀 Run Locally

```bash
# 1. Clone the repository
git clone https://github.com/<your-username>/<your-repo-name>.git
cd <your-repo-name>

# 2. (Optional) create a virtual environment
python -m venv venv
venv\Scripts\activate        # Windows
source venv/bin/activate     # Mac / Linux

# 3. Install requirements
pip install -r requirements.txt

# 4. Run the app
streamlit run app.py
```

The app opens at `http://localhost:8501`.

---

## 🔑 Demo Login (for Judges)

To see the student view, choose **Student (View Own Data)** in the sidebar and sign in:

- **Student ID:** `STU1002`
- **PIN:** `1002` (PIN = last 4 digits of the Student ID)

Other roles (Administrator, Faculty / Mentor, Placement Officer) need no login.

---

## 📁 Project Structure

```
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md           # Project documentation
```

---

## 👥 Team

| Name | Registration No. | Role |
|---|---|---|
| Riya | 251302199 | Team Leader |
| Punita | 251302195 | Team Member |
| Priya Patel | 251302226 | Team Member |
| Yashika Solanki | 251302227 | Team Member |
| Mehak | 251302132 | Team Member |

---

## 📄 License

This project was created for educational and hackathon purposes.
