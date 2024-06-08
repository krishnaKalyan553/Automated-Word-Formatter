
# Automated Word Formatter

A tool that automates the document formatting process for E3s Templates.




## Installation

1) Install the latest stable version of Python
2) clone the repository using 
```
git clone https://github.com/krishnaKalyan553/Automated-Word-Formatter.git
``` 
3) Then install all the required libraries using the command 
```
pip install -r requirements.txt
``` 
4) Start the Django server using the command
```
python manage.py run server
```


## Tech Stack

**Client:** Django Templates

**Document Manipulation:** 
- Docx is used to extract document content and enables us to create a new document object.
- Setting the layout settings for newly created documents before adding the content.  

**Named Entity Recognition (NER):**
- This technique is used to identify the author's name to add stying based on this factor.

**Regex:**
- Used to add conditional formatting rules for elements like Abstract.

**Server:** Django



## Features

- Page Layout Setting
- Heading Non Heading formatting and justificaiton by applying appropriate style
- Author name styling
- Pattern based styking based on regex. 
