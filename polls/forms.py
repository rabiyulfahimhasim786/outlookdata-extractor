from django import forms

class UploadFileForm(forms.Form):
    file = forms.FileField()
# from django import forms
from .models import EmailModel

class EmailDetailsform(forms.ModelForm):
    class Meta:
        model=EmailModel
        fields="__all__"