from flask_wtf import FlaskForm
from wtforms import SubmitField, MultipleFileField
from flask_wtf.file import FileField


class UploadForms(FlaskForm):
    files = MultipleFileField('', render_kw={'multiple': True})
    submit = SubmitField('Submit')