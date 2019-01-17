import tempfile
import time
from pathlib import Path
from typing import Tuple, Sequence
from paramiko import SSHClient, AutoAddPolicy
from scp import SCPClient, SCPException
from PIL import Image

USER = "vendor"
PASSWD = "vendor"


def read_frame_buffer(file: Path) -> Image:
    data = file.read_bytes()
    image = Image.frombytes('RGBA', (1280,800), data)
    b, g, r, a = image.split()
    return Image.merge("RGBA", (r, g, b, a))


def execute_remote_cmd(ssh: SSHClient, command: str) -> Tuple[Sequence[str], Sequence[str]]:
    # Send the command (non-blocking)
    stdin, stdout, stderr = ssh.exec_command(command, get_pty=True)

    # Wait for the command to terminate
    while not stdout.channel.exit_status_ready() and not stdout.channel.recv_ready():
        time.sleep(0.001)

    exit_status = stdout.channel.recv_exit_status()          # Blocking call
    if exit_status != 0:
        raise RuntimeError("Unable to capture remote screen: {}".format(exit_status))
    stdoutstring = stdout.readlines()
    stderrstring = stderr.readlines()
    return stdoutstring, stderrstring


def capture_screen(ssh_ip: str, ssh_port: int) -> Image:
    command = 'su -c "cat /dev/fb0 > /tmp/screen.fb"'
    remote_path = "/tmp/screen.fb"
    local_path = "{}/local_screen.fb".format(tempfile.gettempdir())
    with SSHClient() as ssh_cli:
        ssh_cli.set_missing_host_key_policy(AutoAddPolicy())
        ssh_cli.connect(ssh_ip, port=ssh_port, username=USER, password=PASSWD)
        stdoutstring, stderrstring = execute_remote_cmd(ssh_cli, command)
        with SCPClient(ssh_cli.get_transport()) as scp:
            try:
                scp.get(remote_path, local_path=local_path)
            except SCPException as error:
                raise RuntimeError("Unable to copy remote image: {}".format(error))

    return read_frame_buffer(Path(local_path))
