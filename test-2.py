# coding: utf-8

# In[1]:




import os




exe_path= os.getcwd()
for files in os.walk(exe_path):
    print files
