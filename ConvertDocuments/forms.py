from django import forms

class UploadFileForm(forms.Form):
    title = forms.CharField(max_length=254)
    file = forms.FileField(uploadto = '../../.')