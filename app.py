import os
import msal 
import spacy
import torch
from flask import Flask, render_template, redirect, url_for, session, request
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
import PyPDF2 
import docx 
from sentence_transformers import SentenceTransformer

preprocessed_docs = {}  # Initialize global dictionary

# Initialize Flask and Flask-Login
app = Flask(__name__, template_folder='templates', static_folder='static')
app.secret_key = 'bR7Fj9PzS2Xv5Lq8Mn6Yt3Wk9Qd0Jx4Z'  # Change this to a secure random value in production

login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"


# Office 365 App credentials
CLIENT_ID = "04ca7d78-4284-4358-83f7-53cfbaf720e0"
CLIENT_SECRET = "bcffcc93-7048-4e55-967b-0082bd90f7d6"
AUTHORITY = "https://login.microsoftonline.com/1f8e1b2e-5d9b-44a7-91fc-dbe7d9a51b15"
REDIRECT_URI = "http://localhost:5000/login/callback"
SCOPE = ["User.Read"]


@app.route('/')
def home():
    return redirect(url_for('chat'), code=302)

# Load NLP model
# Initialize spaCy NLP model
nlp = spacy.load("en_core_web_sm")
model = SentenceTransformer("all-MiniLM-L6-v2")

# 🔹 Ensure function definitions come before usage
def extract_text_from_pdf(file_path):
    text = ""
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text += page.extract_text() or ""
    return text

def extract_text_from_docx(file_path):
    text = ""
    doc = docx.Document(file_path)
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

# 🔹 Process documents at startup
def process_documents():
    global preprocessed_docs
    folder = "documents"
    if not os.path.exists(folder):
        print("❌ Documents folder not found!")
        return
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        if filename.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif filename.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        else:
            continue
        sentences = text.lower().split('. ')
        preprocessed_docs[filename] = [sentence.strip() for sentence in sentences]
    print(f"✅ Loaded {len(preprocessed_docs)} documents:", preprocessed_docs.keys())

process_documents()

# User class (Flask-Login)
class User(UserMixin):
    def __init__(self, id, email):
        self.id = id
        self.email = email

# Load user from session
@login_manager.user_loader
def load_user(user_id):
    return User(user_id, session.get("email"))

# Login route
@app.route('/login')
def login():
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri=REDIRECT_URI)
    return redirect(auth_url)

# Callback route
@app.route('/login/callback')
def login_callback():
    msal_app = msal.ConfidentialClientApplication(CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET)
    code = request.args.get('code')
    result = msal_app.acquire_token_by_authorization_code(code, scopes=SCOPE, redirect_uri=REDIRECT_URI)
    if "access_token" in result:
        session["user"] = result.get("id_token_claims")
        session["email"] = session["user"]["preferred_username"]
        user = User(session["user"]["preferred_username"], session["user"]["preferred_username"])
        login_user(user)
        return redirect(url_for('chat'))
    else:
        return "Login failed", 400


@app.route('/chat', methods=['GET', 'POST'])
#@login_required  # Ensures only logged-in users can access chat
def chat():
    print("Current User:", current_user)  # Debugging
    response = None
    if request.method == 'POST':
        user_input = request.form['user_input']
        response = get_answer_from_documents(user_input)
        print("User Input:", user_input)  # Debugging
        print("Response:", response)
    return render_template('chat.html', response=response)

def get_answer_from_documents(query):
    query_embedding = model.encode(query, convert_to_tensor=True)
    ranked_sentences = []
    print(f"🔍 Processing query: {query}")
    print(f"📂 Available documents: {preprocessed_docs.keys()}")
    for doc_name, sentences in preprocessed_docs.items():
        for sentence in sentences:
            sentence_embedding = model.encode(sentence, convert_to_tensor=True)
            similarity_score = torch.nn.functional.cosine_similarity(query_embedding, sentence_embedding, dim=0).item()
            ranked_sentences.append((similarity_score, sentence))
    ranked_sentences.sort(reverse=True, key=lambda x: x[0])
    if ranked_sentences:
        print(f"✅ Top Matches: {ranked_sentences[:3]}")
        return " ".join([s[1] for s in ranked_sentences[:3]])
    print("❌ No relevant information found.")
    return "No relevant information found."

# Logout route
@app.route('/logout')
@login_required
def logout():
    logout_user()
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)

