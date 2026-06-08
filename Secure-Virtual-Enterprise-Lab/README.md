# Secure Virtual Enterprise Lab using VMware ESXi

## Academic Project Documentation

**Course:** DevOps & Cybersecurity Lab
**Project Type:** Enterprise Infrastructure Simulation
**Author:** [Student Name]
**Date:** April 2026
**Version:** 1.0

---

## Project Overview

This project simulates a complete enterprise IT infrastructure using virtualization technology. It demonstrates how organizations can build, configure, and secure their network environment using VMware ESXi as the hypervisor platform. The lab provides hands-on experience with real-world DevOps and cybersecurity practices.

### Purpose

The primary objective is to create a sandboxed environment where students can:
- Practice enterprise network administration
- Learn Active Directory and Windows Server management
- Implement network security using firewalls
- Understand Linux server administration
- Perform security testing and vulnerability assessments
- Implement logging and monitoring solutions

---

## Lab Components

| Component | Operating System | Role |
|-----------|-----------------|------|
| DC-SRV01 | Windows Server 2022 | Domain Controller, DNS, AD |
| WIN-CLI01 | Windows 10/11 | Domain Client Workstation |
| LN-SRV01 | Ubuntu Server 22.04 LTS | Web Server, SSH, Monitoring |
| PF-FW01 | pfSense 2.7 | Firewall, Router, NAT |
| KL-TEST01 | Kali Linux 2024.x | Penetration Testing Machine |

---

## Project Structure

```
Secure-Virtual-Enterprise-Lab/
├── README.md                    # Project overview (this file)
├── docs/
│   ├── project-objectives.md    # Detailed objectives
│   ├── system-architecture.md   # Technical architecture
│   ├── security-features.md    # Security implementation details
│   └── project-report.md       # Final academic report
├── network-diagrams/
│   ├── network-topology.md      # Main network diagram
│   ├── vlan-architecture.md    # VLAN segmentation details
│   └── traffic-flow.md         # Traffic flow analysis
├── configs/
│   ├── pfsense/
│   │   ├── firewall-rules.conf  # pfSense firewall rules
│   │   └── nat-rules.conf       # NAT configuration
│   ├── windows-server/
│   │   ├── gpo-settings.xml     # Group Policy Objects
│   │   └── dns-config.md        # DNS configuration
│   ├── linux-server/
│   │   ├── nginx-config.conf    # Nginx web server config
│   │   └── ssh-hardening.conf   # SSH security configuration
│   └── monitoring/
│       └── log-collection.conf  # Log aggregation setup
├── scripts/
│   ├── powershell/
│   │   ├── ad-setup.ps1         # Active Directory setup script
│   │   ├── user-creation.ps1    # User/group management
│   │   └── gpo-backup.ps1        # GPO backup script
│   ├── bash/
│   │   ├── linux-hardening.sh   # Linux security hardening
│   │   ├── monitoring-setup.sh  # Monitoring installation
│   │   └── firewall-config.sh    # Linux firewall rules
│   └── python/
│       ├── network-scanner.py   # Network scanning utility
│       └── log-analyzer.py      # Log analysis tool
└── setup-guides/
    ├── 01-esxi-host-setup.md    # ESXi installation guide
    ├── 02-vm-creation.md        # Virtual machine creation
    ├── 03-network-setup.md      # vSwitch and VLAN setup
    ├── 04-pfsense-setup.md      # Firewall configuration
    ├── 05-active-directory.md   # AD Domain setup
    ├── 06-linux-server.md       # Linux server setup
    ├── 07-security-testing.md   # Penetration testing guide
    └── 08-monitoring.md         # Monitoring configuration
```

---

## Network Architecture Summary

```
                    ┌─────────────────────────────────────┐
                    │           INTERNET                  │
                    └─────────────────┬───────────────────┘
                                      │
                                      │ WAN
                                      │
                    ┌─────────────────┴───────────────────┐
                    │                                     │
                    │      PF-FW01 (pfSense Firewall)     │
                    │      WAN: DHCP from ISP             │
                    │      LAN: 192.168.10.1              │
                    │      DMZ: 192.168.20.1              │
                    │      MGMT: 192.168.30.1            │
                    └─────────────────┬───────────────────┘
                                      │
          ┌───────────────────────────┼───────────────────┐
          │                           │                   │
          │ LAN Network               │ DMZ Network       │ MGMT Network
          │ 192.168.10.0/24          │ 192.168.20.0/24   │ 192.168.30.0/24
          │                           │                   │
    ┌─────┴─────┐              ┌─────┴─────┐       ┌─────┴─────┐
    │           │              │           │       │           │
┌───┴───┐   ┌───┴───┐      ┌───┴───┐       │   ┌───┴───┐       │
│DC-SRV01│   │WIN-CLI01│      │LN-SRV01│       │   │KL-TEST01│       │
│Domain  │   │Client   │      │Web     │       │   │Pentest  │       │
│Controller│   │Workstation│    │Server │       │   │Machine │       │
│192.168.10.10│192.168.10.20│  │192.168.20.10│   │   │192.168.30.10│   │
└────────┘   └─────────┘      └────────┘       │   └────────┘       │
                                             │                     │
                                        ┌────┴─────┐               │
                                        │ESXi Host │               │
                                        │Management│               │
                                        │192.168.30.100│            │
                                        └──────────┘               │
                                                                   │
```

---

## IP Addressing Scheme

| Network | Subnet | Gateway | DHCP Range | Purpose |
|---------|--------|---------|------------|---------|
| LAN | 192.168.10.0/24 | 192.168.10.1 | 192.168.10.100-200 | Internal Network |
| DMZ | 192.168.20.0/24 | 192.168.20.1 | 192.168.20.100-200 | Public Services |
| MGMT | 192.168.30.0/24 | 192.168.30.1 | 192.168.30.100-200 | Management Network |

### Static IP Assignments

| Hostname | IP Address | Network | Description |
|----------|------------|---------|-------------|
| PF-FW01-LAN | 192.168.10.1 | LAN | pfSense LAN Interface |
| PF-FW01-DMZ | 192.168.20.1 | DMZ | pfSense DMZ Interface |
| PF-FW01-MGMT | 192.168.30.1 | MGMT | pfSense Management Interface |
| DC-SRV01 | 192.168.10.10 | LAN | Domain Controller |
| WIN-CLI01 | 192.168.10.20 | LAN | Windows Client |
| LN-SRV01 | 192.168.20.10 | DMZ | Linux Web Server |
| KL-TEST01 | 192.168.30.10 | MGMT | Kali Linux Testing |
| ESXi-MGMT | 192.168.30.100 | MGMT | ESXi Management Interface |

---

## Quick Start Guide

### Prerequisites

1. **Hardware Requirements:**
   - CPU: Intel VT-x or AMD-V capable processor (4+ cores recommended)
   - RAM: Minimum 16GB (32GB recommended)
   - Storage: 500GB+ SSD
   - Network: 2+ Gigabit Ethernet ports

2. **Software Requirements:**
   - VMware ESXi 7.0 or 8.0
   - VMware vCenter (optional)
   - VMware Workstation/Fusion or vSphere Client for management

3. **ISO Files Required:**
   - Windows Server 2022 Evaluation
   - Windows 10/11 ISO
   - Ubuntu Server 22.04 LTS
   - pfSense 2.7 CE
   - Kali Linux 2024.x

### Installation Order

1. Install ESXi on bare metal (see `setup-guides/01-esxi-host-setup.md`)
2. Create vSwitches and VLANs (see `setup-guides/03-network-setup.md`)
3. Deploy pfSense VM (see `setup-guides/04-pfsense-setup.md`)
4. Deploy Windows Server and configure AD (see `setup-guides/05-active-directory.md`)
5. Deploy Linux server (see `setup-guides/06-linux-server.md`)
6. Deploy Windows client and join domain
7. Deploy Kali Linux for testing (see `setup-guides/07-security-testing.md`)
8. Configure monitoring (see `setup-guides/08-monitoring.md`)

---

## Learning Objectives

Upon completion of this project, students will be able to:

1. **Virtualization**
   - Install and configure VMware ESXi
   - Create and manage virtual machines
   - Configure virtual networking with VLANs

2. **Windows Administration**
   - Deploy Active Directory Domain Services
   - Configure DNS and DHCP services
   - Implement Group Policy Objects
   - Manage users, groups, and organizational units

3. **Linux Administration**
   - Install and configure Ubuntu Server
   - Set up web servers (Nginx/Apache)
   - Implement SSH security hardening
   - Configure firewall rules (UFW/iptables)

4. **Network Security**
   - Configure pfSense firewall
   - Implement network segmentation
   - Create firewall rules for traffic control
   - Configure NAT and port forwarding

5. **Security Testing**
   - Perform network reconnaissance
   - Conduct vulnerability assessments
   - Test firewall rule effectiveness
   - Document security findings

6. **Monitoring & Logging**
   - Centralize log collection
   - Analyze authentication logs
   - Monitor network traffic
   - Implement basic SIEM concepts

---

## Security Features Implemented

| Feature | Implementation | Location |
|---------|---------------|----------|
| Network Segmentation | VLANs via pfSense | Network Layer |
| Firewall | pfSense with rule sets | Perimeter |
| Authentication | Active Directory | Windows Server |
| Access Control | Group Policies | Domain Level |
| SSH Hardening | Key-based auth, limited users | Linux Server |
| Web Security | Nginx SSL/TLS | Linux Server |
| Logging | Centralized collection | All Systems |
| Monitoring | Basic SIEM implementation | LN-SRV01 |

---

## Assessment Criteria

This project will be evaluated based on:

1. **Infrastructure Setup (25%)**
   - Correct VM deployment
   - Network connectivity
   - IP addressing scheme

2. **Active Directory Configuration (20%)**
   - Domain controller setup
   - User/group management
   - Group policy implementation

3. **Firewall & Security (25%)**
   - pfSense configuration
   - Firewall rules
   - Network segmentation

4. **Linux Server (15%)**
   - Web server functionality
   - Security hardening
   - Service configuration

5. **Documentation (15%)**
   - Completeness
   - Clarity
   - Professional presentation

---

## Resources

- [VMware ESXi Documentation](https://docs.vmware.com/en/VMware-vSphere/index.html)
- [pfSense Documentation](https://docs.netgate.com/pfsense/en/latest/)
- [Microsoft Active Directory](https://docs.microsoft.com/en-us/windows-server/identity/ad-ds/active-directory-domain-services)
- [Ubuntu Server Guide](https://ubuntu.com/server/docs)
- [Kali Linux Documentation](https://www.kali.org/docs/)

---

## License

This project is created for educational purposes as part of an academic curriculum.

---

## Disclaimer

This lab environment should only be used for educational purposes. The techniques and tools demonstrated should not be used on production networks or systems without proper authorization. Always practice ethical hacking and follow responsible disclosure guidelines.