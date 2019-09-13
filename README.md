Convert Outlook .msg Files to .eml (MIME format)
================================================

This repository contains a Python 3.6 module for
reading Microsoft Outlook .msg files and converting
them to .eml format, which is the standard MIME
format for email messages.

Install the dependencies with:

    pip3.6 install -r requirements.txt

Then either convert a single file by piping:

	python3.6 outlookmsgfile.py < message.msg > message.eml

Or convert a set of files:

	python3.6 outlookmsgfile.py *.msg

When passing filenames as command-line arguments, a new file with `.eml`
appended to the filename is written out with the message in MIME format.

To use it in your application

    import outlookmsgfile
    eml = outlookmsgfile.load('my_email_sample.msg')
    
The ``load()`` function returns an [EmailMessage](https://docs.python.org/3/library/email.message.html#email.message.EmailMessage) instance.