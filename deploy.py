import paramiko, os, time, re

HOST = 'workai-21.mr-group.ru'
USER = 'userai-21'
PASS = '_4^x-X)w=tbXk8ng-21'
CONTAINER = 'otdelka-signal'
PROJECT = 'otdelka-signal'
LOCAL = r'\\mr.ru\Service\Personal\ignatov_i\Documents\CloudCode\Отделка SIGNAL'

print('Connecting...')
ssh = paramiko.SSHClient()
ssh.set_missing_host_key_policy(paramiko.AutoAddPolicy())
ssh.connect(HOST, username=USER, password=PASS, timeout=30,
            allow_agent=False, look_for_keys=False)

print('Uploading files...')
sftp = ssh.open_sftp()
rd = f'/home/{USER}/{PROJECT}'
try:
    sftp.mkdir(rd)
except:
    pass

for f in ['constructor.html', 'Dockerfile', 'nginx.conf']:
    sftp.put(os.path.join(LOCAL, f), f'{rd}/{f}')
    print(f'  Uploaded: {f}')
sftp.close()

print('Rebuilding container...')
t = ssh.get_transport()
t.set_keepalive(10)
ch = t.open_session(timeout=300)
ch.settimeout(120)
ch.get_pty(width=200)
ch.invoke_shell()
time.sleep(3)
ch.recv(65536)

commands = [
    f'docker stop {CONTAINER} 2>/dev/null; docker rm {CONTAINER} 2>/dev/null; echo "=== Removed old ==="',
    f'cd /home/{USER}/{PROJECT} && docker build -t {PROJECT}:latest . 2>&1 && echo "=== Build OK ==="',
    f'docker run -d --name {CONTAINER} --restart unless-stopped -p 8080:8080 {PROJECT}:latest && echo "=== Started ==="',
    f'sleep 2 && docker logs {CONTAINER} --tail 5 2>&1',
    f'curl -s -o /dev/null -w "HTTP %{{http_code}}" http://127.0.0.1:8080/ && echo ""',
]

for cmd in commands:
    ch.send(cmd + '\n')
    time.sleep(12)
    out = b''
    while ch.recv_ready():
        out += ch.recv(65536)
    clean = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]|\[\?2004[hl]', '',
                   out.decode('utf-8', errors='replace'))
    print(clean.strip())

ch.close()
ssh.close()
print('\nDeploy complete! URL: http://workai-21.mr-group.ru/otdelka-signal/')
