## Python setup on mac

[https://www.freecodecamp.org/news/python-version-on-mac-update/]

`pyenv install 3.10.4`
`pyenv global 3.10.4`

Check installed versions:  
`pyenv versions`

### Edit zshrc to use the pyenv version

set paths:

`echo 'export PYENV_ROOT="$HOME/.pyenv"' >> ~/.zshrc`
`echo 'export PATH="$PYENV_ROOT/bin:$PATH"' >> ~/.zshrc`

set pyenv:
`echo -e 'if command -v pyenv 1>/dev/null 2>&1; then\n eval "$(pyenv init --path)"\n eval "$(pyenv init -)"\nfi' >> ~/.zshrc`
`reset`

check if the version it's updated:
`python --version`

### Links

https://developers.google.com/sheets/api/quickstart/python

https://martin-thoma.com/configuration-files-in-python/

### Node Setup

install

`npm install googleapis --save`
