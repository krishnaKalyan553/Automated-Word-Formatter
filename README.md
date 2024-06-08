Automated Word Formatter
A tool which automates the document formatting process for E3s Templates.

Installation
Install the lateset stable version of Python
clone the repository using
git clone https://github.com/krishnaKalyan553/Automated-Word-Formatter.git
Then install all the required libraries using the command
pip install -r requirements.txt
Start the Django server using the command
python manage.py run server
Tech Stack
Client: Django Templates

Document Manipulation:

Docx is used to extract document content and enables us to create a new document object.
Setting the layout settings for newly created document before adding the content.
Named Entity Recognition (NER):

This technique is used to identify the author name to add stying based on this factor.
Regex:

Used to add conditional formatting rules for elements like Abstract.
Server: Django

Features
Page Layout Setting
Heading Non Heading formatting and justificaiton by applying appropriate style
Author name styling
Pattern based styking based on regex.
