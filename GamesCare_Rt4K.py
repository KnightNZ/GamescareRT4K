import serial
import time
import requests
import socket

def send_serial_command(command, port_name="COM3", baud_rate=115200):
    try:
        with serial.Serial(port_name, baud_rate, timeout=1) as ser:
            ser.write((command + '\n').encode())
            print(f"Executing Command: {command}")
            time.sleep(0.05)  # Small wait between commands
    except Exception as e:
        print(f"Error: {e}")

def send_rt4k(commands, text=None):
    if text:
        print(text)
    
    valid_commands = ["pwr", "menu", "up", "down", "left", "right", "ok", "back", "diag", "stat", "input", "output", "scaler", "sfx", "adc", "prof", "prof1", "prof2", "prof3", "prof4", "prof5", "prof6", "prof7", "prof8", "prof9", "prof10", "prof11", "prof12", "gain", "phase", "pause", "safe", "genlock", "buffer", "res4k", "res1080p", "res1440p", "res480p", "res1", "res2", "res3", "res4", "aux1", "aux2", "aux3", "aux4", "aux5", "aux6", "aux7", "aux8", "pwr on"]
    
    for command in commands:
        if not command:
            print("Error: No hex string provided. Please run with -Input 'HexString'.")
            continue
        
        if command in valid_commands:
            command = f"remote {command}\n"
        else:
            print(f"Unknown command: '{command}'. Please enter a valid command.")
            return
        
        send_serial_command(command)

def switch_input(url, command=None):
    if command:
        send_rt4k([command])
    time.sleep(0.05)
    
    try:
        response = requests.get(url, timeout=2)
        return response.json()
    except Exception as e:
        print(f"Unable to contact switch: {e}")
        exit(1)

def update_labels(games_care_ip):
    try:
        switch_ports = requests.get(f"http://{games_care_ip}/ports", timeout=2).json()
    except Exception as e:
        print(f"Unable to contact switch: {e}")
        exit(1)
    
    ports = switch_ports['Ports']
    active_port = switch_ports['Active']
    
    for i in range(8):
        port_name = ports[i]['Title'] or f"Port {i + 1}"
        port_detected = ports[i]['Detected']
        
        status = ""
        if port_detected == 'True':
            status += ' (*)'
        
        if i + 1 == active_port:
            status += ' (A)'
        
        label = f"{i + 1}: {port_name}{status}"
        

# Main Loop
switch_name = "gcswitch.local"
games_care_ip = None

print("Searching for GamesCare switch on local network...\n")

try:
    games_care_ip = socket.gethostbyname(switch_name)
except Exception as e:
    print("Games Care Switch not found - You may need to specify your IP explicitly.")
    exit(1)

print(f"Games Care Switch found @ {games_care_ip}")

while True:
    update_labels(games_care_ip)
    time.sleep(0.1)
