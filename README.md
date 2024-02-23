Convert Outlook .msg Files to .eml (MIME format)
================================================

This repository contains a Python 3.9+ module for
reading Microsoft Outlook .msg files and converting
them to .eml format, which is the standard MIME
format for email messages.

Install the dependencies with:

    pip install -r requirements.txt

(You may need to create and activate a Python virtual environment first.)

Then either convert a single file by piping:

	python outlookmsgfile.py < message.msg > message.eml

Or convert a set of files:

	python outlookmsgfile.py *.msg

When passing filenames as command-line arguments, a new file with `.eml`
appended to the filename is written out with the message in MIME format.

To use it in your application

    import outlookmsgfile
    eml = outlookmsgfile.load('my_email_sample.msg')
    
The ``load()`` function returns an [EmailMessage](https://docs.python.org/3/library/email.message.html#email.message.EmailMessage) instance.