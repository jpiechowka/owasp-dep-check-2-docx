# owasp-dep-check-2-docx
Parse OWASP Dependency Check reports and create docx document with summary of vulnerabilities

### Requirements
Python-docx is required to run this script (https://python-docx.readthedocs.io/en/latest/). It can be installed by running ```pip install python-docx```

### Usage
1. Create OWASP Dependency Check report in CSV format (https://jeremylong.github.io/DependencyCheck/). ```--format CSV``` command line argument can be used when using CLI.
2. Move report file to ```owasp``` directory alongside the script
3. Run ```python3 ./owasp-dep-check-2-docx.py```
