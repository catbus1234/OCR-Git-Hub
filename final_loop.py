#OCR Project for Advent Energy, converting contact information into excel file
#Author: Josh Chen

#Import necessary packages
from PIL import Image
import pytesseract as tess
import os
import pandas as pd
import re

#Tesseract location
tess.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

#Defining global variables:
#create dataframe to be added to and converted to excel
df = pd.DataFrame(columns = ['name', 'position', 'company name', 'city', 'state','country', 'phone number', 'second phone', 'email'])

#directory of photos
cd = r'C:\Users\Josh\OneDrive - Davidson College\Desktop\OCR_gig\photos2'

#pattern recognition in text
pronoun_pattern = re.compile(r'\((?:\w+/)+\w+\)')
phone_number_pattern = re.compile(r'\(\d{3}\) \d{3}-\d{4}')
email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
state_pattern = re.compile(r'^([A-Za-z\s]+),\s([A-Za-z\s]+)$')

#threshold value for separating background from text
threshold_value = 200

#image coordinates: 750x1334, width x height
tl = 10 #topleft
bl = 250 #bottomleft
tr = 700 #topright
br = 290 #bottomright
big_bound = (10, 500, 740, 1230) #bound for email and office phone numbers


#Converts image to grayscale and defines threshold to separate text from background
def modify_img(img):
    img = img.convert('L') #converts to grayscale
    img = img.point(lambda x: 0 if x < threshold_value else 255)
    return img


#For loop to identify characters in each photo
for photo in os.listdir(cd): #for each photo
    img = Image.open(os.path.join(cd, photo))
    #img = modify_img(img) modify image after scanning for image?

    #Detect if image:
    image_cropped = img.crop((10, 250, 220, 460))
    if tess.image_to_string(image_cropped) == '': #if yes image
        tl = 240 #shift topleft towards right side where there is no image
    bound = (tl, bl, tr, br)


    # Finding the name
    name_text = tess.image_to_string(img.crop(bound))

    #detect if pronouns
    pronouns_bound = (tl, 250, 680, 350)
    pronoun_matches = pronoun_pattern.findall(tess.image_to_string(img.crop(pronouns_bound)))

    if(len(pronoun_matches) > 0):
        bl = 330
        br = 370
    else:
        bl = 300
        br = 340
    bound = (tl, bl, tr, br)
        
    #Find the position
    position_text = tess.image_to_string(img.crop(bound))
    bl += 50
    br += 50
    bound = (tl, bl, tr, br) #update bounds

    #Finding the company
    company_text = tess.image_to_string(img.crop(bound))
    bl += 50
    br += 50
    bound = (tl, bl, tr, br) #update bounds

    #Finding the city
    text = tess.image_to_string(img.crop(bound))
    match = re.match(state_pattern, text)
    #Detect if there is text for city, state
    city_text = match.group(1) if match else text
    state_text = match.group(2) if match else ''
    bl += 50
    br += 50
    bound = (tl, bl, tr, br) #update bounds

    #Finding the country
    country_text = tess.image_to_string(img.crop(bound))
    if len(country_text) == 0: #if no country labeled: fill field with United States
        country_text = "United States"

    #Finding phone numbers
    text = tess.image_to_string(img.crop(big_bound))
    phone_matches = phone_number_pattern.findall(text)

    #Determining how many phone numbers to record
    first_phone = phone_matches[0] if len(phone_matches) > 0 else ''
    second_phone = phone_matches[1] if len(phone_matches) > 1 else ''

    #Finding email
    email_matches = email_pattern.findall(text)
    email = email_matches[0] if len(email_matches) > 0 else ''

    # Create a dictionary with the data from the current photo
    data = {
        'name': name_text,
        'position': position_text,
        'company name': company_text,
        'city': city_text,
        'state': state_text,
        'country': country_text,
        'phone number': first_phone,
        'second phone': second_phone,  
        'email': email
    }

    # Append the data to the DataFrame
    df.loc[len(df)] = data

    #resetting coordinates
    tl = 10 #topleft
    bl = 250 #bottomleft
    tr = 700 #topright
    br = 290 #bottomright
    bound = (tl, bl, tr, br)


df.to_excel('database2.xlsx', index=False)

