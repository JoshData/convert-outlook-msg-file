Convert Outlook .msg Files to .eml (MIME format)
================================================

This repository contains a Python 3.6 module for
reading Microsoft Outlook .msg files and converting
them to .eml format, which is the standard MIME
format for email messages.

The module requires Python 3.6 and the [compoundfiles](https://pypi.python.org/pypi/compoundfiles)
package, so first install that:

	pip3.6 install compoundfiles

We also rely on our Python 3 port of [compressed_rtf](https://github.com/delimitry/compressed_rtf), which is included in this repository.

Then either convert a single file by piping:

	python3.6 outlookmsgfile.py < message.msg > message.eml

Or convert a set of files:

	python3.6 outlookmsgfile.py *.msg

When passing filenames as command-line arguments, a new file with `.eml`
appended to the filename is written out with the message in MIME format.
