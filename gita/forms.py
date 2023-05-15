from django import forms
import datetime


class ImageUploadForm(forms.Form):
    photo = forms.ImageField(label = 'PP-size Photo', widget=forms.FileInput(attrs={'class': 'form-control-file'}), required=False)  # photo
    sign = forms.ImageField(label = 'Signature', widget=forms.FileInput(attrs={'class': 'form-control-file'}), required=False) # Signature
    eng_name = forms.CharField(label = 'Full Name', widget=forms.TextInput(attrs={'class': 'form-control form-control-sm', 'placeholder': 'In English. E.g. Ram Bahadur Subedi'}))
    nep_name = forms.CharField(label = 'पुरा नाम', widget=forms.TextInput(attrs={'oninput': 'validateDevanagari(this)', 'class': 'form-control form-control-sm', 'placeholder': 'In Nepali. E.g. राम बहादुर सुवेदी'}))
    dOB = forms.DateTimeField(
        label='Date of Birth',
        widget=forms.SelectDateWidget(
            years=range(1925, 2015),
            attrs={'class': 'form-control'}
        ),
        initial=datetime.datetime(2000, 1, 1)  # Set the initial value to January 1, 2000
    )
    yearCHOICES = [(i, str(i)) for i in range(1, 6)]
    semCHOICES = [(i, str(i)) for i in range(1, 3)]
    year = forms.ChoiceField(label = 'Year', choices=yearCHOICES, widget=forms.Select(attrs={'class': 'form-check form-check-inline'}))
    sem = forms.ChoiceField(label = 'Semester', choices=semCHOICES, widget=forms.Select(attrs={'class': 'form-check form-check-inline'}))
    genderCHOICES = [('female', 'Female'), ('male', 'Male')]
    genders = forms.ChoiceField(label='Gender', choices=genderCHOICES, widget=forms.Select)

