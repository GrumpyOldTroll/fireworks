# Overview

This is a script to help translate a fireworks cue sheet into the layout of the racks of fireworks launchers, to assist with planning.

I don't have a sample input file I can publish, but it's apparently a semi-standard format (though I don't know a specification for it?) consisting of several columns detailing the information about the fireworks.  It's used as the cue sheet for launching the fireworks in a show.  We used this to help plan the setup for the 2024 Torrance fireworks show.

This program ignores most of the columns but looks for a header with column names, and want it to include:

 * PIN: which firing pin the fireworks should be attached to
 * QTY: the number of shells of this caliber attached to the pin
 * CAL: the caliber of the shells.  Should be 101 for a 4-inch shell, 76 for a 3-inch shell, and 63 for a 2.5-inch shell.

The idea is that shells of the same caliber go together with other shells in a rack of usually 3x5 or 4x5 and placed nearby their firing strip, in roughly the order they need to be fired.  3" shell racks alternate with 4" shell racks, with counts based on what's needed in the cue sheet.

To set up the environment, installing the requirements in a [venv](https://docs.python.org/3/library/venv.html) is recommended:

~~~
python3 -m venv fwshow
source fwshow/bin/activate
pip install -r requirements.in
~~~

Running (assuming you're inside the vm, by having run `source fwshow/bin/activate`):

~~~
python main.py cue_sheet.xlsx
~~~

or for the reversed-phase layout, starting with 4" instead of 3":

~~~
python main.py --phased cue_sheet.xlsx
~~~

Running this with a valid input file should generate either `fireworks_boards.xlsx` or `fireworks_boards_flipped.xlsx`, which you can import into google sheets by dragging.

The resulting sheet should have 2 tabs:
 - "Kim Slave": shows how many shells of each size in each rack get attached to the corresponding board
 - "Kim Slave Layout": shows the specific locations and pins of a corresponding arrangement of shells.

TBD: write up how this is used coherently with pictures, this is gonna make no sense to anyone except a few select people atm...

