# DCOM Lateral Movement – Black Hat MEA 2023

> **Presented by:** Faisal Sultan Alabduljabbar  
> **Event:** Black Hat MEA 2023, Riyadh  
> **Topic:** Advanced Machine Pivoting Using DCOM  
> **Twitter:** [@Fsalabduljabbar](https://x.com/Fsalabduljabbar)  
> **LinkedIn:** [faisal-alabduljabbar](https://www.linkedin.com/in/faisal-alabduljabbar-%F0%9F%87%B8%F0%9F%87%A6-24309013a/)

---

## 🧠 Summary

This repository includes a full **proof-of-concept (PoC)** that demonstrates how adversaries can abuse **Distributed Component Object Model (DCOM)** for stealthy **lateral movement** across Windows environments. The DCOM payload functions similarly to an SMB payload.

The code was demonstrated during my session at **Black Hat MEA 2023**.

---

## ⚙️ Components

### ✅ COM Server (Out-of-Proc EXE)

- Hosts a COM interface exposed via DCOM on port `135`.
- Accepts requests from remote clients to execute commands (e.g., `systeminfo`).
- Writes output to disk (`C:\Windows\Temp\logs.txt`).

### ✅ COM Client

- Connects to a remote machine using DCOM.
- Authenticates using provided credentials.
- Calls `COMSend(...)` to request tasks and send results.

---

## 🚀 Usage

### Server Side (on Target Machine)

1. Register the COM server using `regsvr32` or by embedding registration logic (`SetupRegistryKeys()`).
2. Make sure port **135** is accessible.
3. Place `COMServer.exe` in an accessible location.

### Client Side (Attacker System)

```bash
COMClient.exe <TargetIP> <Username> <Password> <Domain>
```

The client will:

- Connect to the remote DCOM server.
- Ask for a command (default: `systeminfo`).
- Execute the command.
- Send the result back.

---

## ⚠️ Disclaimer

This repository is for educational and authorized security research only. Do **not** use these techniques on any system without explicit permission.

Feel free to contribute improvements or questions.  
For collaboration or consulting, contact me via [LinkedIn](https://www.linkedin.com/in/faisal-alabduljabbar-%F0%9F%87%B8%F0%9F%87%A6-24309013a/).
