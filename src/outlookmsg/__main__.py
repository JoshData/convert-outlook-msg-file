#! /usr/bin/env python

# This module converts a Microsoft Outlook .msg file into
# a MIME message that can be loaded by most email programs
# or inspected in a text editor.
#
# This script relies on the Python package compoundfiles
# for reading the .msg container format.
#
# References:
#
# https://msdn.microsoft.com/en-us/library/cc463912.aspx
# https://msdn.microsoft.com/en-us/library/cc463900(v=exchg.80).aspx
# https://msdn.microsoft.com/en-us/library/ee157583(v=exchg.80).aspx
# https://blogs.msdn.microsoft.com/openspecification/2009/11/06/msg-file-format-part-1/

import logging
import sys

from .stream import load

logger = logging.getLogger(__name__)

# COMMAND-LINE ENTRY POINT


def main():
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s - %(name)s:%(funcName)s:%(lineno)s - %(levelname)s - %(message)s",
    )
    # If no command-line arguments are given, convert the .msg
    # file on STDIN to .eml format on STDOUT.
    if len(sys.argv) <= 1:
        print(load(sys.stdin), file=sys.stdout)

    # Otherwise, for each file mentioned on the command-line,
    # convert it and save it to a file with ".eml" appended
    # to the name.
    else:
        for fn in sys.argv[1:]:
            print(fn + "...")
            msg = load(fn)
            with open(fn + ".eml", "wb") as f:
                f.write(msg.as_bytes())


if __name__ == "__main__":
    main()
