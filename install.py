import subprocess, sys

def system(cmd):
    """Run system command cmd."""
    failure, output = subprocess.getstatusoutput(cmd)
    if failure:
        print(output)
        sys.exit(1)
    else:
        print(output)

system('python3 -m ensurepip')
system("pip3 --version")
system("pip3 install -r requirements.txt")