import io
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
try:
    with open('out.txt', encoding='utf-16-le') as f:
        for line in f:
            if 'Ведущий инженер по промышленной' in line:
                print(line.strip())
except Exception as e:
    print(e)
