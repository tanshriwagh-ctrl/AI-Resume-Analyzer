from flask import Flask, request, jsonify
from flask_cors import CORS
import mysql.connector
from werkzeug.security import generate_password_hash, check_password_hash
from werkzeug.utils import secure_filename
from pypdf import PdfReader
from docx import Document
import os
import re
import uuid

app = Flask(__name__)
CORS(app)

# =========================================================
# SETTINGS
# =========================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")

ALLOWED_EXTENSIONS = {"pdf", "docx"}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Allows multiple resumes in one upload request.
app.config["MAX_CONTENT_LENGTH"] = 600 * 1024 * 1024


# =========================================================
# MYSQL DATABASE
# =========================================================
def get_db_connection():
    return mysql.connector.connect(
        host="localhost",
        user="root",
        password="",
        database="ai_resume_analyzer"
    )


# =========================================================
# BASIC HELPERS
# =========================================================
def allowed_file(filename):
    return (
        "." in filename
        and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS
    )


def normalize_text(text):
    text = text or ""
    text = text.lower()
    text = text.replace("–", "-").replace("—", "-")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def contains_term(text, term):
    text = normalize_text(text)
    term = normalize_text(term)

    if not term:
        return False

    # Flexible matching for punctuation used in names such as
    # React.js, Node.js, C++, C#, .NET, etc.
    escaped = re.escape(term)
    escaped = escaped.replace(r"\ ", r"\s+")
    escaped = escaped.replace(r"\.", r"\.?")

    return re.search(
        r"(?<![a-z0-9])" + escaped + r"(?![a-z0-9])",
        text
    ) is not None


# =========================================================
# RESUME TEXT EXTRACTION
# =========================================================
def extract_resume_text(file_path):
    extension = os.path.splitext(file_path)[1].lower()

    if extension == ".pdf":
        reader = PdfReader(file_path)
        pages = []

        for page in reader.pages:
            pages.append(page.extract_text() or "")

        return "\n".join(pages)

    if extension == ".docx":
        document = Document(file_path)
        parts = []

        for paragraph in document.paragraphs:
            if paragraph.text.strip():
                parts.append(paragraph.text)

        for table in document.tables:
            for row in table.rows:
                parts.append(" ".join(cell.text for cell in row.cells))

        return "\n".join(parts)

    return ""


# =========================================================
# SKILL NORMALIZATION
# =========================================================
# Skills are grouped so that common equivalent spellings are
# treated as the same skill.
SKILL_ALIASES = {
    "python": {"python", "python3"},
    "javascript": {"javascript", "js", "ecmascript"},
    "typescript": {"typescript", "ts"},
    "react": {"react", "react.js", "reactjs"},
    "angular": {"angular", "angular.js", "angularjs"},
    "vue": {"vue", "vue.js", "vuejs"},
    "node.js": {"node", "node.js", "nodejs"},
    "express": {"express", "express.js", "expressjs"},
    "flask": {"flask"},
    "django": {"django"},
    "php": {"php"},
    "laravel": {"laravel"},
    "java": {"java"},
    "spring": {"spring", "spring boot", "springboot"},
    "c++": {"c++", "cpp"},
    "c#": {"c#", "c sharp", "csharp"},
    ".net": {".net", "dotnet", "asp.net", "asp net"},
    "sql": {"sql", "mysql", "postgresql", "postgres", "sqlite"},
    "mysql": {"mysql"},
    "postgresql": {"postgresql", "postgres"},
    "mongodb": {"mongodb", "mongo"},
    "oracle": {"oracle"},
    "html": {"html", "html5"},
    "css": {"css", "css3"},
    "bootstrap": {"bootstrap"},
    "tailwind": {"tailwind", "tailwind css"},
    "rest api": {"rest api", "restful api", "restful services", "api development"},
    "git": {"git", "github", "gitlab", "version control"},
    "github": {"github"},
    "docker": {"docker", "containerization"},
    "kubernetes": {"kubernetes", "k8s"},
    "aws": {"aws", "amazon web services"},
    "azure": {"azure", "microsoft azure"},
    "gcp": {"gcp", "google cloud", "google cloud platform"},
    "machine learning": {"machine learning", "ml"},
    "deep learning": {"deep learning", "dl"},
    "data analysis": {"data analysis", "data analytics"},
    "pandas": {"pandas"},
    "numpy": {"numpy"},
    "tensorflow": {"tensorflow"},
    "pytorch": {"pytorch"},
    "power bi": {"power bi", "powerbi"},
    "excel": {"excel", "microsoft excel"},
    "figma": {"figma"},
}


def canonical_skill(skill):
    value = normalize_text(skill)

    for canonical, aliases in SKILL_ALIASES.items():
        if value == canonical or value in {normalize_text(a) for a in aliases}:
            return canonical

    return value


def skill_is_present(resume_text, required_skill):
    skill = normalize_text(required_skill)
    canonical = canonical_skill(skill)

    aliases = SKILL_ALIASES.get(
        canonical,
        {skill}
    )

    for alias in aliases:
        if contains_term(resume_text, alias):
            return True

    return False


def split_required_skills(required_skills_text):
    if isinstance(required_skills_text, list):
        raw = required_skills_text
    else:
        raw = re.split(r"[,;\n]+", str(required_skills_text or ""))

    skills = []

    for item in raw:
        item = str(item).strip()
        if item and item.lower() not in [x.lower() for x in skills]:
            skills.append(item)

    return skills


# =========================================================
# EXPERIENCE EXTRACTION
# =========================================================
def get_required_years(job_description):
    text = normalize_text(job_description)

    patterns = [
        r"(\d+(?:\.\d+)?)\s*\+?\s*(?:years?|yrs?)\s+(?:of\s+)?experience",
        r"(\d+(?:\.\d+)?)\s*\+?\s*(?:years?|yrs?)"
    ]

    values = []

    for pattern in patterns:
        for value in re.findall(pattern, text):
            try:
                values.append(float(value))
            except ValueError:
                pass

        if values:
            break

    return max(values) if values else 0


def get_resume_years(text):
    normalized = normalize_text(text)
    values = []

    # Example: "2 years experience"
    for value in re.findall(
        r"(\d+(?:\.\d+)?)\s*\+?\s*(?:years?|yrs?)\s+(?:of\s+)?experience",
        normalized
    ):
        try:
            values.append(float(value))
        except ValueError:
            pass

    # Example: "2019 - 2023", "2022 - Present"
    ranges = re.findall(
        r"\b(19\d{2}|20\d{2})\s*-\s*(19\d{2}|20\d{2}|present|current)\b",
        normalized
    )

    for start, end in ranges:
        try:
            start_year = int(start)

            if end in ("present", "current"):
                end_year = 2026
            else:
                end_year = int(end)

            if end_year >= start_year:
                values.append(float(end_year - start_year))
        except ValueError:
            pass

    return max(values) if values else 0


def has_experience_information(resume_text):
    return any(
        contains_term(resume_text, term)
        for term in [
            "experience",
            "work experience",
            "professional experience",
            "employment",
            "internship",
            "intern"
        ]
    )


# =========================================================
# EDUCATION
# =========================================================
EDUCATION_GROUPS = {
    "bca": {"bca", "bachelor of computer applications"},
    "bba": {"bba", "bachelor of business administration"},
    "btech": {"b.tech", "btech", "bachelor of technology"},
    "be": {"b.e", "be", "bachelor of engineering"},
    "bsc": {"b.sc", "bsc", "bachelor of science"},
    "mca": {"mca", "master of computer applications"},
    "mba": {"mba", "master of business administration"},
    "mtech": {"m.tech", "mtech", "master of technology"},
    "msc": {"m.sc", "msc", "master of science"},
    "computer science": {"computer science"},
    "computer applications": {"computer applications"},
    "information technology": {"information technology", "information technology engineering", "it"},
}


def education_matches(resume_text, job_description):
    resume = normalize_text(resume_text)
    job = normalize_text(job_description)

    required_groups = []

    for group, terms in EDUCATION_GROUPS.items():
        if any(contains_term(job, term) for term in terms):
            required_groups.append(group)

    if required_groups:
        for group in required_groups:
            if any(contains_term(resume, term) for term in EDUCATION_GROUPS[group]):
                return True, True

        # Resume has education but not the specifically requested degree.
        has_any_education = any(
            contains_term(resume, term)
            for terms in EDUCATION_GROUPS.values()
            for term in terms
        )
        return False, has_any_education

    # If job does not specify a degree, give credit for having
    # a recognizable education section/qualification.
    has_education = any(
        contains_term(resume, term)
        for term in [
            "education",
            "bachelor",
            "master",
            "degree",
            "bca",
            "bba",
            "b.tech",
            "btech",
            "b.e",
            "bsc",
            "mca",
            "mba",
            "mtech",
            "msc"
        ]
    )

    return has_education, has_education


# =========================================================
# CERTIFICATIONS
# =========================================================
CERTIFICATION_TERMS = [
    "certification",
    "certifications",
    "certified",
    "certificate",
    "aws certified",
    "azure certification",
    "google cloud certification",
    "microsoft certified",
    "oracle certified",
    "coursera certificate",
    "udemy certificate"
]


def certification_match(resume_text, job_description):
    resume = normalize_text(resume_text)
    job = normalize_text(job_description)

    # First check whether the job specifically asks for a
    # certification/provider.
    requested = []

    for term in [
        "aws",
        "azure",
        "google cloud",
        "gcp",
        "oracle",
        "microsoft certified",
        "certification",
        "certified"
    ]:
        if contains_term(job, term):
            requested.append(term)

    if requested:
        for term in requested:
            if contains_term(resume, term):
                return True, True

        has_any_cert = any(
            contains_term(resume, term)
            for term in CERTIFICATION_TERMS
        )
        return False, has_any_cert

    has_any_cert = any(
        contains_term(resume, term)
        for term in CERTIFICATION_TERMS
    )

    return has_any_cert, has_any_cert


# =========================================================
# JOB KEYWORDS
# =========================================================
STOPWORDS = {
    "the", "and", "for", "with", "that", "this", "from",
    "have", "has", "are", "you", "your", "our", "will",
    "should", "must", "candidate", "candidates", "required",
    "preferred", "good", "knowledge", "work", "working",
    "using", "into", "their", "they", "about", "also",
    "years", "year", "role", "roles", "job", "jobs",
    "developer", "development", "team", "teams", "skills",
    "skill", "responsibilities", "requirements", "requirement",
    "looking", "seeking", "ability", "strong", "including",
    "including", "such", "etc", "least", "plus"
}


def extract_keywords(job_title, required_skills_text, job_description):
    source = " ".join([
        str(job_title or ""),
        str(required_skills_text or ""),
        str(job_description or "")
    ])

    words = re.findall(
        r"[a-zA-Z][a-zA-Z0-9+#.-]{2,}",
        source.lower()
    )

    keywords = []

    # Required skills have higher priority.
    for skill in split_required_skills(required_skills_text):
        skill = normalize_text(skill)
        if skill and skill not in keywords:
            keywords.append(skill)

    # Then add useful words from the job title/description.
    for word in words:
        word = word.strip(".,;:()[]{}")
        if (
            word
            and word not in STOPWORDS
            and word not in keywords
            and len(word) >= 3
        ):
            keywords.append(word)

    return keywords[:50]


# =========================================================
# ATS SCORE CALCULATION
# =========================================================
def calculate_analysis(
    resume_text,
    job_title,
    required_skills_text,
    job_description
):
    resume_original = resume_text or ""
    resume = normalize_text(resume_original)
    job = normalize_text(job_description)

    required_skills = split_required_skills(required_skills_text)

    # ---------------------------------------------------------
    # 1. SKILLS MATCH / 40
    # ---------------------------------------------------------
    matched_skills = []
    missing_skills = []

    for skill in required_skills:
        if skill_is_present(resume, skill):
            matched_skills.append(skill)
        else:
            missing_skills.append(skill)

    if required_skills:
        skills_score = round(
            (len(matched_skills) / len(required_skills)) * 40
        )
    else:
        skills_score = 0

    # ---------------------------------------------------------
    # 2. EXPERIENCE / 20
    # ---------------------------------------------------------
    required_years = get_required_years(job_description)
    resume_years = get_resume_years(resume_original)

    if required_years > 0:
        if resume_years >= required_years:
            experience_score = 20
        elif resume_years > 0:
            experience_score = round(
                (resume_years / required_years) * 20
            )
        elif has_experience_information(resume):
            experience_score = 5
        else:
            experience_score = 0
    else:
        experience_score = 20 if has_experience_information(resume) else 0

    # ---------------------------------------------------------
    # 3. EDUCATION / 15
    # ---------------------------------------------------------
    exact_education, has_any_education = education_matches(
        resume,
        job_description
    )

    if exact_education:
        education_score = 15
    elif has_any_education:
        education_score = 7
    else:
        education_score = 0

    # ---------------------------------------------------------
    # 4. CERTIFICATIONS / 10
    # ---------------------------------------------------------
    exact_certification, has_any_certification = certification_match(
        resume,
        job_description
    )

    if exact_certification:
        certification_score = 10
    elif has_any_certification:
        certification_score = 5
    else:
        certification_score = 0

    # ---------------------------------------------------------
    # 5. KEYWORDS & FORMAT / 15
    # ---------------------------------------------------------
    keywords = extract_keywords(
        job_title,
        required_skills_text,
        job_description
    )

    keyword_matches = []

    for keyword in keywords:
        # Required skills are already scored separately, so the
        # keyword section focuses more on job-description context.
        if contains_term(resume, keyword):
            keyword_matches.append(keyword)

    if keywords:
        keyword_score = round(
            (len(keyword_matches) / len(keywords)) * 12
        )
    else:
        keyword_score = 0

    format_score = 0

    if len(resume_original.strip()) >= 300:
        format_score += 1

    if re.search(
        r"\b[\w.+-]+@[\w.-]+\.[a-zA-Z]{2,}\b",
        resume_original
    ):
        format_score += 1

    section_count = sum(
        1
        for section in [
            "skills",
            "education",
            "experience",
            "projects",
            "certifications"
        ]
        if contains_term(resume, section)
    )

    if section_count >= 3:
        format_score += 1

    keywords_score = min(
        keyword_score + format_score,
        15
    )

    # ---------------------------------------------------------
    # FINAL SCORE
    # ---------------------------------------------------------
    total_score = min(
        skills_score
        + experience_score
        + education_score
        + certification_score
        + keywords_score,
        100
    )

    # ---------------------------------------------------------
    # MATCH LEVEL
    # ---------------------------------------------------------
    if total_score >= 80:
        match_level = "Excellent Match"
    elif total_score >= 65:
        match_level = "Good Match"
    elif total_score >= 50:
        match_level = "Moderate Match"
    else:
        match_level = "Low Match"

    # ---------------------------------------------------------
    # STRENGTHS
    # ---------------------------------------------------------
    strengths = []

    if matched_skills:
        strengths.append(
            "Matched skills: " + ", ".join(matched_skills)
        )

    if required_skills and not missing_skills:
        strengths.append("All required skills were detected")

    if required_years > 0 and resume_years >= required_years:
        strengths.append(
            f"Meets the required experience of about {required_years:g} years"
        )
    elif required_years == 0 and experience_score >= 15:
        strengths.append("Relevant experience information found")

    if exact_education:
        strengths.append("Education matches the job requirement")

    if exact_certification:
        strengths.append("Relevant certification requirement matched")

    if not strengths:
        strengths.append("Limited relevant information was detected")

    # ---------------------------------------------------------
    # WEAKNESSES
    # ---------------------------------------------------------
    weaknesses = []

    if missing_skills:
        weaknesses.append(
            "Missing skills: " + ", ".join(missing_skills)
        )

    if required_years > 0 and resume_years < required_years:
        weaknesses.append(
            f"Required experience: about {required_years:g} years; "
            f"detected experience: about {resume_years:g} years"
        )

    if job and not exact_education and has_any_education:
        weaknesses.append(
            "Education is present but does not clearly match the specified requirement"
        )
    elif job and not has_any_education:
        weaknesses.append("Education information was not clearly detected")

    if not exact_certification and not has_any_certification:
        # Only mention this as a weakness if the job asks for certification.
        if any(
            contains_term(job, term)
            for term in ["certification", "certified", "aws", "azure", "gcp", "google cloud"]
        ):
            weaknesses.append("Required certification was not detected")

    if not weaknesses:
        weaknesses.append("No major mismatch detected by the ATS rules")

    # ---------------------------------------------------------
    # SUGGESTIONS
    # ---------------------------------------------------------
    suggestions = []

    if missing_skills:
        suggestions.append(
            "Add missing required skills only if the candidate genuinely has them."
        )

    if required_years > 0 and resume_years < required_years:
        suggestions.append(
            "Clearly highlight relevant projects, internships and work experience."
        )

    if not exact_education and not has_any_education:
        suggestions.append(
            "Add clear education details such as degree and specialization."
        )

    if not has_any_certification:
        suggestions.append(
            "Mention relevant certifications if the candidate has them."
        )

    suggestions.append(
        "Keep Skills, Experience, Education, Projects and Certifications in clear sections."
    )

    return {
        "ats_score": total_score,
        "match_level": match_level,
        "skills": skills_score,
        "experience": experience_score,
        "education": education_score,
        "certifications": certification_score,
        "keywords": keywords_score,
        "matched_skills": matched_skills,
        "missing_skills": missing_skills,
        "matched_keywords": keyword_matches[:30],
        "strengths": "; ".join(strengths),
        "weaknesses": "; ".join(weaknesses),
        "suggestions": " ".join(suggestions)
    }


# =========================================================
# HOME
# =========================================================
@app.route("/")
def home():
    return "AI Resume Analyzer Backend is Running!"


# =========================================================
# TEST DATABASE
# =========================================================
@app.route("/test-db")
def test_db():
    try:
        db = get_db_connection()
        cursor = db.cursor()

        cursor.execute("SELECT COUNT(*) FROM users")
        result = cursor.fetchone()

        cursor.close()
        db.close()

        return jsonify({
            "status": "success",
            "message": "MySQL database connected successfully!",
            "users_count": result[0]
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


# =========================================================
# REGISTER
# =========================================================
@app.route("/register", methods=["POST"])
def register():
    try:
        data = request.get_json() or {}

        name = str(data.get("name", "")).strip()
        email = str(data.get("email", "")).strip()
        password = data.get("password", "")

        if not name or not email or not password:
            return jsonify({
                "status": "error",
                "message": "Please fill all required fields."
            }), 400

        db = get_db_connection()
        cursor = db.cursor()

        cursor.execute(
            "SELECT user_id FROM users WHERE email = %s",
            (email,)
        )

        if cursor.fetchone():
            cursor.close()
            db.close()

            return jsonify({
                "status": "error",
                "message": "Email already registered."
            }), 409

        hashed_password = generate_password_hash(password)

        cursor.execute(
            """
            INSERT INTO users (name, email, password)
            VALUES (%s, %s, %s)
            """,
            (name, email, hashed_password)
        )

        db.commit()

        cursor.close()
        db.close()

        return jsonify({
            "status": "success",
            "message": "Registration successful!"
        }), 201

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


# =========================================================
# LOGIN
# =========================================================
@app.route("/login", methods=["POST"])
def login():
    try:
        data = request.get_json() or {}

        email = str(data.get("email", "")).strip()
        password = data.get("password", "")

        if not email or not password:
            return jsonify({
                "status": "error",
                "message": "Please enter email and password."
            }), 400

        db = get_db_connection()
        cursor = db.cursor(dictionary=True)

        cursor.execute(
            """
            SELECT user_id, name, email, password
            FROM users
            WHERE email = %s
            """,
            (email,)
        )

        user = cursor.fetchone()

        cursor.close()
        db.close()

        if not user or not check_password_hash(user["password"], password):
            return jsonify({
                "status": "error",
                "message": "Invalid email or password."
            }), 401

        return jsonify({
            "status": "success",
            "message": "Login successful!",
            "user": {
                "user_id": user["user_id"],
                "name": user["name"],
                "email": user["email"]
            }
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


# =========================================================
# SAVE JOB REQUIREMENTS
# =========================================================
@app.route("/jobs", methods=["POST"])
def create_job():
    try:
        data = request.get_json() or {}

        user_id = data.get("user_id")
        job_title = str(data.get("job_title", "")).strip()

        required_skills = data.get("required_skills", "")

        # job.html sends an array. Convert it to one database string.
        if isinstance(required_skills, list):
            required_skills = ", ".join(
                str(skill).strip()
                for skill in required_skills
                if str(skill).strip()
            )
        else:
            required_skills = str(required_skills).strip()

        job_description = str(
            data.get("job_description", "")
        ).strip()

        if not user_id:
            return jsonify({
                "status": "error",
                "message": "Please login first."
            }), 401

        if not job_title:
            return jsonify({
                "status": "error",
                "message": "Please enter job title."
            }), 400

        if not required_skills:
            return jsonify({
                "status": "error",
                "message": "Please add at least one required skill."
            }), 400

        if not job_description:
            return jsonify({
                "status": "error",
                "message": "Please enter job description."
            }), 400

        db = get_db_connection()
        cursor = db.cursor()

        cursor.execute(
            "SELECT user_id FROM users WHERE user_id = %s",
            (user_id,)
        )

        if not cursor.fetchone():
            cursor.close()
            db.close()

            return jsonify({
                "status": "error",
                "message": "User not found. Please login again."
            }), 401

        cursor.execute(
            """
            INSERT INTO jobs
            (user_id, job_title, required_skills, job_description)
            VALUES (%s, %s, %s, %s)
            """,
            (
                user_id,
                job_title,
                required_skills,
                job_description
            )
        )

        db.commit()

        job_id = cursor.lastrowid

        cursor.close()
        db.close()

        return jsonify({
            "status": "success",
            "message": "Job requirements saved successfully!",
            "job": {
                "job_id": job_id,
                "user_id": user_id,
                "job_title": job_title,
                "required_skills": required_skills,
                "job_description": job_description
            }
        }), 201

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


# =========================================================
# GET JOBS
# =========================================================
@app.route("/jobs/<int:user_id>", methods=["GET"])
def get_jobs(user_id):
    try:
        db = get_db_connection()
        cursor = db.cursor(dictionary=True)

        cursor.execute(
            """
            SELECT
                job_id,
                user_id,
                job_title,
                required_skills,
                job_description,
                created_at
            FROM jobs
            WHERE user_id = %s
            ORDER BY job_id DESC
            """,
            (user_id,)
        )

        jobs = cursor.fetchall()

        cursor.close()
        db.close()

        return jsonify({
            "status": "success",
            "jobs": jobs
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


# =========================================================
# UPLOAD RESUMES
# =========================================================
@app.route("/upload-resumes", methods=["POST"])
def upload_resumes():
    try:
        user_id = request.form.get("user_id")

        if not user_id:
            return jsonify({
                "status": "error",
                "message": "Please login first."
            }), 401

        files = request.files.getlist("resumes")

        if not files:
            return jsonify({
                "status": "error",
                "message": "Please select at least one resume."
            }), 400

        db = get_db_connection()
        cursor = db.cursor()

        uploaded = []

        for file in files:
            if not file or not file.filename:
                continue

            if not allowed_file(file.filename):
                continue

            original_name = secure_filename(file.filename)

            extension = os.path.splitext(
                original_name
            )[1].lower()

            stored_name = (
                f"{uuid.uuid4().hex}{extension}"
            )

            full_path = os.path.join(
                UPLOAD_FOLDER,
                stored_name
            )

            file.save(full_path)

            db_path = os.path.join(
                "backend",
                "uploads",
                stored_name
            ).replace("\\", "/")

            cursor.execute(
                """
                INSERT INTO resumes
                (user_id, file_name, file_path)
                VALUES (%s, %s, %s)
                """,
                (
                    user_id,
                    original_name,
                    db_path
                )
            )

            uploaded.append({
                "resume_id": cursor.lastrowid,
                "file_name": original_name
            })

        if not uploaded:
            cursor.close()
            db.close()

            return jsonify({
                "status": "error",
                "message": "No valid PDF or DOCX resumes were uploaded."
            }), 400

        db.commit()

        cursor.close()
        db.close()

        return jsonify({
            "status": "success",
            "message": f"{len(uploaded)} resume(s) uploaded successfully!",
            "resumes": uploaded
        }), 201

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


# =========================================================
# GET RESUMES
# =========================================================
@app.route("/resumes/<int:user_id>", methods=["GET"])
def get_resumes(user_id):
    try:
        db = get_db_connection()
        cursor = db.cursor(dictionary=True)

        cursor.execute(
            """
            SELECT
                resume_id,
                user_id,
                file_name,
                file_path,
                uploaded_at
            FROM resumes
            WHERE user_id = %s
            ORDER BY resume_id DESC
            """,
            (user_id,)
        )

        resumes = cursor.fetchall()

        cursor.close()
        db.close()

        return jsonify({
            "status": "success",
            "resumes": resumes
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


# =========================================================
# ANALYZE RESUMES
# =========================================================
@app.route("/analyze-resumes", methods=["POST"])
def analyze_resumes():
    try:
        data = request.get_json() or {}

        user_id = data.get("user_id")
        job_id = data.get("job_id")
        resume_ids = data.get("resume_ids", [])

        if not user_id:
            return jsonify({
                "status": "error",
                "message": "Please login first."
            }), 401

        if not job_id:
            return jsonify({
                "status": "error",
                "message": "Job requirement not found."
            }), 400

        if not resume_ids:
            return jsonify({
                "status": "error",
                "message": "No resumes selected for analysis."
            }), 400

        # Make sure IDs are integers.
        try:
            resume_ids = [int(x) for x in resume_ids]
        except (TypeError, ValueError):
            return jsonify({
                "status": "error",
                "message": "Invalid resume IDs."
            }), 400

        db = get_db_connection()
        cursor = db.cursor(dictionary=True)

        # -----------------------------------------------------
        # GET THE EXACT JOB SELECTED BY THE RECRUITER
        # -----------------------------------------------------
        cursor.execute(
            """
            SELECT
                job_id,
                job_title,
                required_skills,
                job_description
            FROM jobs
            WHERE job_id = %s
              AND user_id = %s
            """,
            (job_id, user_id)
        )

        job = cursor.fetchone()

        if not job:
            cursor.close()
            db.close()

            return jsonify({
                "status": "error",
                "message": "Selected job requirement was not found."
            }), 404

        # -----------------------------------------------------
        # GET ONLY THE SELECTED RESUMES
        # -----------------------------------------------------
        placeholders = ",".join(
            ["%s"] * len(resume_ids)
        )

        query = f"""
            SELECT
                resume_id,
                file_name,
                file_path
            FROM resumes
            WHERE user_id = %s
              AND resume_id IN ({placeholders})
        """

        cursor.execute(
            query,
            [user_id] + resume_ids
        )

        resumes = cursor.fetchall()

        if not resumes:
            cursor.close()
            db.close()

            return jsonify({
                "status": "error",
                "message": "Selected resumes were not found."
            }), 404

        results = []

        # -----------------------------------------------------
        # ANALYZE EVERY SELECTED RESUME
        # -----------------------------------------------------
        for resume in resumes:
            stored_path = resume["file_path"]

            # Normal stored path:
            # backend/uploads/filename.pdf
            filename_only = os.path.basename(stored_path)

            possible_paths = [
                os.path.join(BASE_DIR, "uploads", filename_only),
                os.path.join(
                    BASE_DIR,
                    stored_path.replace("/", os.sep)
                ),
            ]

            full_path = None

            for possible in possible_paths:
                if os.path.exists(possible):
                    full_path = possible
                    break

            # -------------------------------------------------
            # FILE NOT FOUND
            # -------------------------------------------------
            if not full_path:
                missing = split_required_skills(
                    job["required_skills"]
                )

                results.append({
                    "resume_id": resume["resume_id"],
                    "file_name": resume["file_name"],
                    "ats_score": 0,
                    "match_level": "Low Match",
                    "skills": 0,
                    "experience": 0,
                    "education": 0,
                    "certifications": 0,
                    "keywords": 0,
                    "matched_skills": [],
                    "missing_skills": missing,
                    "matched_keywords": [],
                    "strengths": "Resume file could not be found.",
                    "weaknesses": "Stored resume file is missing.",
                    "suggestions": "Upload the resume again."
                })

                continue

            # -------------------------------------------------
            # EXTRACT RESUME TEXT
            # -------------------------------------------------
            try:
                resume_text = extract_resume_text(full_path)
            except Exception as file_error:
                results.append({
                    "resume_id": resume["resume_id"],
                    "file_name": resume["file_name"],
                    "ats_score": 0,
                    "match_level": "Low Match",
                    "skills": 0,
                    "experience": 0,
                    "education": 0,
                    "certifications": 0,
                    "keywords": 0,
                    "matched_skills": [],
                    "missing_skills": [],
                    "matched_keywords": [],
                    "strengths": "",
                    "weaknesses": "Could not read resume: " + str(file_error),
                    "suggestions": "Check that the PDF/DOCX file is readable."
                })

                continue

            # Empty/scanned PDF protection.
            if not resume_text.strip():
                results.append({
                    "resume_id": resume["resume_id"],
                    "file_name": resume["file_name"],
                    "ats_score": 0,
                    "match_level": "Low Match",
                    "skills": 0,
                    "experience": 0,
                    "education": 0,
                    "certifications": 0,
                    "keywords": 0,
                    "matched_skills": [],
                    "missing_skills": split_required_skills(
                        job["required_skills"]
                    ),
                    "matched_keywords": [],
                    "strengths": "No readable text was extracted from the resume.",
                    "weaknesses": "The resume may be scanned/image-based or unreadable.",
                    "suggestions": "Use a text-based PDF/DOCX resume and upload it again."
                })

                continue

            # -------------------------------------------------
            # REAL JOB-SPECIFIC ATS CALCULATION
            # -------------------------------------------------
            result = calculate_analysis(
                resume_text,
                job["job_title"],
                job["required_skills"],
                job["job_description"]
            )

            # -------------------------------------------------
            # SAVE ANALYSIS RESULT
            # -------------------------------------------------
            cursor2 = db.cursor()

            cursor2.execute(
                """
                INSERT INTO analysis_results
                (
                    resume_id,
                    ats_score,
                    strengths,
                    weaknesses,
                    suggestions
                )
                VALUES (%s, %s, %s, %s, %s)
                """,
                (
                    resume["resume_id"],
                    result["ats_score"],
                    result["strengths"],
                    result["weaknesses"],
                    result["suggestions"]
                )
            )

            cursor2.close()

            result["resume_id"] = resume["resume_id"]
            result["file_name"] = resume["file_name"]

            results.append(result)

        db.commit()

        cursor.close()
        db.close()

        # -----------------------------------------------------
        # HIGHEST MATCH FIRST
        # -----------------------------------------------------
        results.sort(
            key=lambda item: item["ats_score"],
            reverse=True
        )

        # Add rank after sorting.
        for index, item in enumerate(results, start=1):
            item["rank"] = index

        return jsonify({
            "status": "success",
            "message": "Resume analysis completed successfully!",
            "job": {
                "job_id": job["job_id"],
                "job_title": job["job_title"],
                "required_skills": split_required_skills(
                    job["required_skills"]
                )
            },
            "results": results
        })

    except Exception as e:
        return jsonify({
            "status": "error",
            "message": str(e)
        }), 500


# =========================================================
# RUN SERVER
# =========================================================
if __name__ == "__main__":
    app.run(debug=True)
