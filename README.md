# EMS01
Exam Management Software. 

Guys here we present a complete exam management software. It can simply by installed by running the following commands on Ubuntu 14.04 terminal having Python 2.7
***********

1. Install Tkinter:

 $ sudo apt-get install python-tk


2. Installing openpyxl:

   First install the 'jdcal' package bt the following command:

     $ sudo pip install jdcal

   Download openpyxl package from https://pypi.python.org/packages/source/o/openpyxl/openpyxl-2.3.4.tar.gz (Simply paste the url   in your browser.) Save the file in Downloads folder.

   Now open terminal and type

     $ cd ~/Downloads

     $ tar -xvf openpyxl-2.3.4.tar.gz

     $ cd openpyxl-2.3.4/

     $ sudo python setup.py install


3. Install MongoDB:

   1.  $ sudo apt-key adv --keyserver hkp://keyserver.ubuntu.com:80 --recv 7F0CEB10

   2.  $ echo 'deb http://downloads-distro.mongodb.org/repo/ubuntu-upstart dist 10gen' | sudo tee /etc/apt/sources.list.d/mongodb.list

   3.  $ sudo apt-get update

   4.  $ sudo apt-get install -y mongodb-org



4. Install pymongo:

   $ sudo apt-get install python-pymongo

5. Install Image-Tk:

   $ sudo apt-get install python-imaging-tk

Download the source code file home.py.
Also download the image BG.png and examhelp.pdf and keep them in same directory as of home.py.
You can even download sample input files for the softwares from the same repository and check the execution.

Configuring for your pc:

Go to line number 480 in file home.py and replace "/home/vijay/Codes/Python/" by the directory address where home.py and examhelp.pdf are located.

Then execute the python code by the command: $  python home.py

Hope you will enjoy it!!! :)

--Den
********
