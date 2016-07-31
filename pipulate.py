
# coding: utf-8

# In[1]:

"""Functions for plugging into Pipulate-frameworks for conducting SEO investigations."""


# In[2]:

import requests, re


# In[3]:

test_url = 'http://mikelev.in/'


# In[4]:

def Title(**row_dict):
    compiled_pattern = re.compile(r'<title\s?>(.*?)</title\s?>', re.DOTALL)
    a_string = row_dict['response'].text
    matches = compiled_pattern.findall(a_string)
    if matches:
        return(200, matches[0].strip())
    else:
        return(200, None)

