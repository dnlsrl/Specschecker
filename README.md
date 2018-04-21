# Specschecker 2

This is a complete rewrite of the [original version](https://gist.github.com/dnlsrl/a286d202481f24e4652ad29899fefb5c), which I created while I was still unexperienced in programming with Python. It's not like I am an expert now, but I know better. If you compare both versions, you will notice that this one is better written (hopefully).

Every suggestion is welcomed, just bear in mind that I might need to research a little before being able to implement something, depending on the difficulty. PRs are also welcome! I just ask not to over clutter the script and to document it properly. Speaking of which, I tried my best to comment as frecuently and as clearly as possible so as for somebody to get the gist of the code with just a look! If you got some feedback for me in that regard, please leave me a comment as well.

### Features:

* Print values to screen
* Export to CSV
* Command line arguments
    * `python specschecker.py -s` will print the data to screen, will save to file via interactive wizard
    * `python specschecker.py -f filename` will save the data to a .csv file, no data printed to screen

### Dependencies:

The script depends on the following libraries:

* [regex](https://pypi.python.org/pypi/regex)
* [pywin32](https://sourceforge.net/projects/pywin32/files/pywin32/)
* [psutil](https://github.com/giampaolo/psutil)

You don't have to worry about them, though, just `install pipenv` via pip, and then, in the project directory run `pipenv install` and pipenv will install the dependencies in a virtual environment. To know more about pipenv, refer to https://docs.pipenv.org/

### Contact

**Email**: [dnlsrl.kaiser@gmail.com](mailto:dnlsrl.kaiser@gmail.com)  
**XMPP/Jabber**: dnlsrl@xmpp.is