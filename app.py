import os
import msal 
import spacy
from flask import Flask, render_template, redirect, url_for, session, request
from flask_login import LoginManager, UserMixin, login_user, login_required, logout_user, current_user
import PyPDF2 
import docx 

# Initialize Flask and Flask-Login
app = Flask(__name__)
app.secret_key = os.urandom(24)
login_manager = LoginManager()
login_manager.init_app(app)
login_manager.login_view = "login"

# Office 365 App credentials
CLIENT_ID = "04ca7d78-4284-4358-83f7-53cfbaf720e0"
CLIENT_SECRET = "bcffcc93-7048-4e55-967b-0082bd90f7d6"
AUTHORITY = "https://login.microsoftonline.com/1f8e1b2e-5d9b-44a7-91fc-dbe7d9a51b15"
REDIRECT_URI = "http://localhost:5000/login/callback"
SCOPE = ["User.Read"]

# Admin user email (only admin can upload documents)
ADMIN_EMAIL = "pramod.pawar@talentica.com"  # Replace with your admin's Office 365 email

# Initialize spaCy NLP model
nlp = spacy.load("en_core_web_sm")

# Define function to extract text from PDFs
def extract_text_from_pdf(file_path):
    text = ""
    with open(file_path, "rb") as file:
        reader = PyPDF2.PdfReader(file)
        for page in range(len(reader.pages)):
            text += reader.pages[page].extract_text()
    return text

# Define function to extract text from DOCX files
def extract_text_from_docx(file_path):
    text = ""
    doc = docx.Document(file_path)
    for para in doc.paragraphs:
        text += para.text + "\n"
    return text

# Define function to process documents and store them in memory
def process_documents():
    documents = []
    for filename in os.listdir("documents"):
        file_path = os.path.join("documents", filename)
        if filename.endswith('.pdf'):
            text = extract_text_from_pdf(file_path)
        elif filename.endswith('.docx'):
            text = extract_text_from_docx(file_path)
        else:
            continue
        documents.append({"filename": filename, "text": text})
    return documents

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

# Chat route (requires login)
@app.route('/chat', methods=['GET', 'POST'])
@login_required
def chat():
    if request.method == 'POST':
        user_input = request.form['user_input']
        documents = process_documents()  # Get the documents from local storage
        response = get_answer_from_documents(user_input, documents)
        return render_template('chat.html', response=response, user_input=user_input)
    return render_template('chat.html', response=None)

# Get answer from documents using NLP (simple approach)
def get_answer_from_documents(query, documents):
    # Process the query with spaCy NLP
    query_doc = nlp(query.lower())
    
    # Find the document that is most relevant to the query
    best_match = None
    best_score = 0
    for document in documents:
        document_text = document["text"].lower()
        score = document_text.count(query.lower())
        
        if score > best_score:
            best_score = score
            best_match = document["filename"]

    if best_match:
        return f"Found answer in: {best_match}"
    else:
        return "Sorry, I couldn't find any relevant information."

# Logout route
@app.route('/logout')
@login_required
def logout():
    logout_user()
    session.clear()
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)
