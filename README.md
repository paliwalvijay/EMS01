# EMS01
Exam Management Software. 

Guys here we present a complete exam management software. It can simply be installed by running the following commands on Ubuntu 14.04 terminal having Python 2.7

Install Tkinter:
sudo apt-get install python-tk

Install openpyxl:
sudo apt-get install python-openpyxl

Install MongoDB:
1. sudo apt-key adv --keyserver hkp://keyserver.ubuntu.com:80 --recv 7F0CEB10
2. echo 'deb http://downloads-distro.mongodb.org/repo/ubuntu-upstart dist 10gen' | sudo tee /etc/apt/sources.list.d/mongodb.list
3. sudo apt-get update
4. sudo apt-get install -y mongodb-org


Install pymongo:
sudo apt-get install python-pymongo

Install Image-Tk:
sudo apt-get install python-imaging-tk

Download the source code file home.py.
Also download the image BG.png and examhelp.pdf and keep them in same directory as of home.py.
You can even download sample input files for the softwares from the same repository and check the execution.

Configuring for your pc:
Go to line number 463 in file home.py and replace "/home/vijay/Codes/Python/" by the directory address where home.py and examhelp.pdf are located.

Then execute the python code by the command: $  python home.py

Hope you will enjoy it!!! :)

--Den
