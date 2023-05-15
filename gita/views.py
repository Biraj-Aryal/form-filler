from django.http import FileResponse, HttpResponse 
from django.shortcuts import render, redirect
import os
import regex
import datetime
from .forms import ImageUploadForm
from .fill import fill_my_form
from django.conf import settings

base_dir = settings.BASE_DIR
uploaded_path = os.path.join(base_dir, 'fill_files')
print('this is the pwd', base_dir)

# Create your views here.

def data_upload(request):
    if request.method == 'POST':
        form = ImageUploadForm(request.POST, request.FILES)
        if form.is_valid():
            handle_uploaded_files(request.FILES.get('photo', 0), 
            request.FILES.get('sign', 0), 
            form.cleaned_data['eng_name'],
            form.cleaned_data['nep_name'],
            form.cleaned_data['year'],
            form.cleaned_data['sem'],
            form.cleaned_data['dOB'].date(),
            form.cleaned_data.get('genders'),
            request.POST.get('numeric-hyphen-field', ''))        

            regis = request.POST.get('numeric-hyphen-field', '')
            docx_file_path = os.path.join(uploaded_path, f"{regis}.docx")
            file_path = docx_file_path
            with open(file_path,'rb') as doc:
                response = HttpResponse(doc.read(), content_type='application/ms-word')
                response['Content-Disposition'] = f'attachment;filename={regis}.docx'
                os.unlink(docx_file_path)
                return response
            return redirect('landing')        



    else:
        form = ImageUploadForm()
    return render(request, 'gita/fill.html', {'form': form})


def handle_uploaded_files(photo, sign, eng_name, nep_name, year, sem, dob, gender, reg_no ):

    print("It is still working!")
    
    # Save the two images
    image1_name = 'image1.jpg'
    image1_path = os.path.join(uploaded_path, image1_name)

    if photo:
        with open(image1_path, 'wb') as destination1:
            for chunk in photo.chunks():
                destination1.write(chunk)

    image2_name = 'image2.jpg'
    image2_path = os.path.join(uploaded_path, image2_name)

    if sign: # checking if user that uploaded a sign image in the form
        with open(image2_path, 'wb') as destination2:
            for chunk in sign.chunks():
                destination2.write(chunk)

    # Save the text content as a text file
    text1_name = f'{reg_no}.txt'
    text1_path = os.path.join(uploaded_path, 'usage', text1_name)
    with open(text1_path, 'w') as text_file:
        text_file.write(str(eng_name)+'\n')
        text_file.write(str(nep_name)+'\n')
        text_file.write(str(dob)+'\n')
        text_file.write(str(gender)+'\n')
        text_file.write(str(reg_no)+'\n')
        text_file.write(str(year)+'\n')
        text_file.write(str(sem))
    
    # APP WORK HERE ONWARDS!
    if photo:
        photo_ = 'image1.jpg'
    else: 
        photo_ = 'empty_photo.png'
    if sign:
        sign_ = 'image2.jpg'
    else: 
        sign_ = 'empty_sign.png'
    eng_name_ = eng_name
    nep_name_ = nep_name
    regis = reg_no
    usr_year = int(year)
    usr_sem = int(sem)
    year_e = str(dob.year)
    month_e = str(dob.month)
    day_e = str(dob.day)
    usr_gender = gender # options: 'male', 'female', 'any_string_really'
    fill_my_form(base_dir, photo_, sign_, regis, usr_year, usr_sem, eng_name_, nep_name_, year_e, month_e, day_e, usr_gender)

    # delete the photos
    if photo: # checking if the user had uploaded the photo in the form
        os.unlink(image1_path)
    if sign:
        os.unlink(image2_path)










