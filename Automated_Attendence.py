#!/usr/bin/env python
# coding: utf-8

# In[1]:


'''
These are the  libraries imported
1- cv2 to manage photos and web cam
2- face_recognition to encode faces and compare it with the training set
3- numpy to manage lists
4- os to manage paths
5- openpyxl to create the attendence excel sheet'''
import cv2
import face_recognition as fr
import numpy as np
import os
import openpyxl 
from openpyxl import Workbook


# In[2]:


'''
lines from:
1-2: Takes the path of the images from my computer and make a list of their names
4-end: I made a new list then put the image names inside of it without the extension'''
path = r"C:\Users\Amr\Desktop\trainingImg"
img_names = os.listdir(path)
print("num of images detected are: ", img_names)
classNames = []
for name in img_names:
    s = name.split(".")[0]
    classNames.append(s)
print(classNames)


# In[3]:


'''
Here I use the image names to complete the path required to load the photos to 
the program then make a list of those photos 
'''
img = []
for imgs in img_names:
    img.append(fr.load_image_file(f"{path}/{imgs}"))
print(len(img))


# In[4]:


'''
I create a list in which the encodings of the photos done by the face recognition 
library are stored'''
encodings = []
for imgs in img:
    encodings.append(fr.face_encodings(imgs)[0])
print(len(encodings))
print("Done encoding")


# In[ ]:


'''
1-4: I made a set and a list to put the names inside of the people who attend 
to the excel sheet 
5: activate the web cam to capture faces
while loop:
I used while as an infinite loop to iterate because sometimes the library gets
it wrong and doesn't recognise a face which will ruin a definite loop, and even
if i didn't iterate when face is not recognized, if someone is absent it will 
still iterate forever
I could: 
for i in range(0,len(encodings-1)), and at the end add: else: i--

2-6 in while loop: I make the web cam capture images and list them in 
an array for cv2 to be able to resize and set colors

7-9: i detect the face location with fr and then gets the face encodings
10-11: just to show the cam for screenshots :) 

12:end: here I compare the encoded images with the encodings of the cam face 
and show the distance between the them (the less the distance the more close
it is

thank you :) have a nice day :) <3 '''
here_names = set()
here_list = []
results_file = Workbook()
results_sheet = results_file.active
cam = cv2.VideoCapture(0)
while True:
    #fisrt line to help put the names in the excel file
    ele_in_set = len(here_names)
    success, cam_img = cam.read()
    cam_img = np.array(cam_img)
    img_small = cv2.resize(cam_img,(0,0),None,0.25,0.25)
    img_small = cv2.cvtColor(cam_img, cv2.COLOR_BGR2RGB)
    face_in_cam = fr.face_locations(img_small)
    encode_cam = fr.face_encodings(img_small, face_in_cam)
    cv2.imshow("cam", cam_img)  
    cv2.waitKey(1)
    for encoded, locations in zip(encode_cam, face_in_cam):
        match = fr.compare_faces(encodings, encoded)
        face_dist = fr.face_distance(encodings, encoded)
        print(face_dist)
        matched_pic = np.argmin(face_dist)
        if match[matched_pic]: 
            print(classNames[matched_pic])
            #adding the name to the excel sheet
            here_names.add(classNames[matched_pic])
            if len(here_names)>ele_in_set:
                here_list.append(classNames[matched_pic])
                results_sheet[f"A{len(here_list)}"] = here_list[len(here_list)-1]
                results_file.save(filename="attendence.xlsx")
                
                





