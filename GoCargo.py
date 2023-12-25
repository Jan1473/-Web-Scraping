#!/usr/bin/env python
# coding: utf-8

# In[60]:


get_ipython().system('pip install selenium')


# In[61]:


import selenium


# In[62]:


from selenium import webdriver


# In[63]:


driver = webdriver.Chrome()


# In[64]:


driver.get('https://www.google.com/search?q=internshala')


# In[65]:


from selenium.webdriver.common.by import By


# In[66]:


f = open('google.csv', 'w' , encoding='utf-8')
f.write('Title,Link,Detail')
for element in driver.find_elements(By.XPATH,'//div[@id="search"]//div[@class="ULSxyf"]'):
    title = element.find_element(By.XPATH, './/h3').text
    link = element.find_element(By.XPATH, './/a[@jsname="UWckNb"]').get_attribute('href')
    detail = element.find_element(By.XPATH ,'//div[@class="VwiC3b yXK7lf lyLwlc yDYNvb W8l4ac lEBKkf"]').text
    f.write(title.replace(',','|')+','+link.replace(',','|')+','+detail.replace(',','|')+'\n')
f.close()


# In[ ]:




