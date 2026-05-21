import urllib.request
import sys

try:
    resp = urllib.request.urlopen('http://127.0.0.1:8000/api/', timeout=5)
    print(resp.read().decode())
except Exception as e:
    print('ERR', e)
    sys.exit(1)
