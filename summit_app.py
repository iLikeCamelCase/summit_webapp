from flask import Flask, render_template, url_for, request, redirect, flash, \
session, send_file
from flask_sqlalchemy import SQLAlchemy
from forms import UploadForms
from werkzeug.utils import secure_filename
import flask_wtf
from flask_bootstrap import Bootstrap
from summit_script.main import Script_Instance

#sveltd

app = Flask(__name__)
bootstrap = Bootstrap(app)

# secrets.token_hex(n)
config = {
    'SECRET_KEY' : '328dfe1ff00345c170e25aec89942741',
    'WTF_CSRF_SECRET_KEY' : '8c2ab876dffe702b7872726d67d9dcdb'
}
app.config.update(config)

UPLOAD_FOLDER = 'summit_script/paystubs/'
ALLOWED_EXTENSIONS = {'pdf'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

@app.route('/', methods=['GET','POST'])
@app.route('/home', methods=['GET','POST'])
def index(show_form=1):
    form = UploadForms()
    #if form.validate_on_submit():
    if 'show_form' not in session:
        session['show_form'] = 1
    if request.method == 'POST':
        
        for file in form.files.data:
            if allowed_file(file.filename):
                file_filename = secure_filename(file.filename)
                file.save(UPLOAD_FOLDER + file_filename)

        # NEED METHOD WHICH VALIDATES FILES AS PDF
        # NEED METHOD WHICH VALIDATES FILES AS PAYSTUBS
        flash('Files uploaded successfully', 'success')
        #return redirect(url_for('index'))
        session['show_form'] = 0
        script = Script_Instance()
        return send_file('summit_script/output/paystub.xlsx', as_attachment=True)
        
    return render_template('index.html', form=form)

@app.route('/update-session')
def update_session():
       session['show_form'] = 1  # Update the session variable
       return redirect(url_for('index'))
    

def allowed_file(filename):
    return '.' in filename and \
    filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


if __name__ == '__main__':
    app.run(debug = True)