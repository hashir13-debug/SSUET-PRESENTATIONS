# SECURE VIRTUAL ENTERPRISE LAB USING VMWARE ESXi

## An Academic Project Report

---


**Student Name:** ________________________________

**Course Name:** ________________________________

**Instructor Name:** ________________________________

**Submission Date:** ________________________________

---

## 1. Introduction

This project demonstrates the design and implementation of a secure virtual enterprise lab environment using VMware ESXi. The lab simulates a real-world corporate IT infrastructure with network segmentation, firewall protection, active directory authentication, and security monitoring. It provides hands-on experience with virtualization technologies, network security, and system administration practices essential for modern enterprise environments.

---

## 2. Objectives

The main objectives of this project are:

- Design and deploy a virtualized IT infrastructure using VMware ESXi
- Implement network segmentation with VLANs for security isolation
- Configure Active Directory for centralized authentication and management
- Deploy and secure a Linux-based web server in a DMZ
- Implement firewall rules to control traffic between network segments
- Perform security testing to validate the security posture
- Establish monitoring and logging for security analysis
- Document the complete implementation process

---

## 3. Tools & Technologies Used

| Category | Technology | Purpose |
|----------|------------|---------|
| Virtualization | VMware ESXi 7.0/8.0 | Hypervisor platform |
| Firewall | pfSense 2.7 | Network security and routing |
| Directory Service | Windows Server 2022 | Active Directory Domain Controller |
| Client OS | Windows 10/11 | Domain-joined workstation |
| Web Server | Ubuntu Server 22.04 | Linux-based web server |
| Security Testing | Kali Linux 2024 | Penetration testing platform |

---

## 4. System Architecture

The system architecture follows a defense-in-depth security model with multiple layers of protection. The infrastructure is segmented into four distinct network zones: WAN (Internet), LAN (Internal Network), DMZ (Public Services), and Management Network. Each zone is isolated by the pfSense firewall, which controls all traffic between segments.

The Domain Controller provides centralized authentication for all Windows systems through Active Directory. The Linux server hosts public-facing web services in the DMZ, isolated from the internal LAN. The management network contains administrative tools and security testing systems.

---

## 5. Network Design

### Network Topology Diagram

```
                          INTERNET
                             |
                             | WAN
                             |
                    +-------------------+
                    |   pfSense Firewall|
                    |    (PF-FW01)      |
                    +-------------------+
                    | WAN | LAN | DMZ | MGMT
                    +-------------------+
                        |    |     |     |
         +--------------+    |     |     +--------------+
         |                   |     |                    |
         | LAN               | DMZ |                    | MGMT
         | VLAN 10           | VLAN 20                 | VLAN 30
         | 192.168.10.0/24   | 192.168.20.0/24        | 192.168.30.0/24
         |                   |                         |
    +----------+        +----------+              +-----------+
    | DC-SRV01 |        | LN-SRV01 |              | KL-TEST01 |
    | Windows  |        | Linux    |              | Kali Linux|
    | Server   |        | Web      |              | Security  |
    | 192.168.10.10     | 192.168.20.10           | 192.168.30.10
    +----------+        +----------+              +-----------+
         |
    +----------+
    |WIN-CLI01 |
    | Windows  |
    | Client   |
    |192.168.10.20
    +----------+
```

### IP Addressing Scheme

| Network | VLAN | Subnet | Gateway | Purpose |
|---------|------|--------|---------|---------|
| WAN | N/A | DHCP | ISP | Internet connection |
| LAN | 10 | 192.168.10.0/24 | 192.168.10.1 | Internal network |
| DMZ | 20 | 192.168.20.0/24 | 192.168.20.1 | Public servers |
| Management | 30 | 192.168.30.0/24 | 192.168.30.1 | Administration |

### Static IP Assignments

| Device | IP Address | Network | Role |
|--------|------------|---------|------|
| pfSense LAN | 192.168.10.1 | LAN | Gateway |
| pfSense DMZ | 192.168.20.1 | DMZ | Gateway |
| pfSense MGMT | 192.168.30.1 | MGMT | Gateway |
| DC-SRV01 | 192.168.10.10 | LAN | Domain Controller |
| WIN-CLI01 | 192.168.10.20 | LAN | Client Workstation |
| LN-SRV01 | 192.168.20.10 | DMZ | Web Server |
| KL-TEST01 | 192.168.30.10 | MGMT | Security Testing |
| ESXi Host | 192.168.30.100 | MGMT | Hypervisor |

---

## 6. Virtual Machines and Roles

### VM Configuration Summary

| VM Name | Operating System | vCPU | RAM | Disk | Network |
|---------|-------------------|------|-----|------|---------|
| PF-FW01 | pfSense 2.7 | 2 | 4 GB | 20 GB | WAN + LAN + DMZ + MGMT |
| DC-SRV01 | Windows Server 2022 | 2 | 4 GB | 60 GB | LAN |
| WIN-CLI01 | Windows 10/11 | 2 | 4 GB | 40 GB | LAN |
| LN-SRV01 | Ubuntu Server 22.04 | 2 | 2 GB | 30 GB | DMZ |
| KL-TEST01 | Kali Linux 2024 | 2 | 4 GB | 50 GB | MGMT |

### VM Roles Description

**PF-FW01 (pfSense Firewall):** Acts as the perimeter security device, routing traffic between all network segments. Implements firewall rules, NAT, and network segmentation.

**DC-SRV01 (Domain Controller):** Runs Active Directory Domain Services, DNS, and DHCP for the internal network. Provides centralized authentication and policy management.

**WIN-CLI01 (Windows Client):** Domain-joined workstation for user access and testing group policy enforcement.

**LN-SRV01 (Linux Server):** Web server located in the DMZ for public-facing services. Hardened with firewall rules, SSH key authentication, and monitoring.

**KL-TEST01 (Kali Linux):** Security testing platform for network scanning, vulnerability assessment, and firewall rule verification.

---

## 7. Active Directory Setup

### Domain Configuration

| Setting | Value |
|---------|-------|
| Domain Name | smitlab.local |
| Domain NetBIOS | SMITLAB |
| Forest Functional Level | Windows Server 2016 |
| Domain Functional Level | Windows Server 2016 |

### Organizational Unit Structure

```
smitlab.local
├── Users
│   ├── Admin_Users
│   └── Standard_Users
├── Computers
│   ├── Workstations
│   └── Servers
└── Groups
    └── Security_Groups
```

### User Accounts Created

| Username | Group | Description |
|----------|-------|-------------|
| itadmin | Domain Admins, SG_IT_Admins | IT Administrator |
| jsmith | SG_StandardUsers | Standard User |
| jdoe | SG_StandardUsers | Standard User |
| testuser | SG_StandardUsers | Test Account |

### Security Groups Created

| Group Name | Purpose |
|------------|---------|
| SG_IT_Admins | IT Administrators with elevated access |
| SG_HelpDesk | Help desk support staff |
| SG_Developers | Development team members |
| SG_StandardUsers | Standard domain users |

### Password Policy Settings

| Policy | Setting |
|--------|---------|
| Minimum Password Length | 12 characters |
| Password Complexity | Enabled |
| Maximum Password Age | 90 days |
| Password History | 24 passwords |
| Account Lockout Threshold | 5 failed attempts |
| Account Lockout Duration | 30 minutes |

---

## 8. Firewall Configuration

### pfSense Interface Configuration

| Interface | IP Address | Network | Purpose |
|-----------|------------|---------|---------|
| WAN | DHCP | Internet | External connection |
| LAN | 192.168.10.1 | LAN | Internal network gateway |
| DMZ | 192.168.20.1 | DMZ | Public services gateway |
| MGMT | 192.168.30.1 | MGMT | Management gateway |

### Firewall Rules Summary

| Rule | Source | Destination | Action | Purpose |
|------|--------|-------------|--------|---------|
| LAN-Internet | LAN | WAN | Allow | Allow users internet access |
| LAN-DMZ | LAN | DMZ | Allow | Allow access to web server |
| LAN-MGMT | LAN | MGMT | Block | Prevent user access to management |
| DMZ-LAN | DMZ | LAN | Block | Critical: Isolate DMZ from LAN |
| DMZ-Internet | DMZ | WAN | Allow | Allow server updates |
| MGMT-All | MGMT | Any | Allow | Allow admin access to all networks |
| WAN-DMZ-HTTP | WAN | DMZ:80,443 | Allow | Allow public web access |

### NAT Configuration

Network Address Translation (NAT) is configured on the WAN interface to allow all internal networks to access the internet using the single public IP address. Outbound NAT rules automatically translate internal IP addresses to the WAN interface address for internet-bound traffic.

---

## 9. Linux Server Setup

### System Configuration Commands

```bash
# Update system packages
sudo apt update
sudo apt upgrade -y

# Set hostname
sudo hostnamectl set-hostname lnsrv01

# Configure static IP
sudo nano /etc/netplan/00-installer-config.yaml

# Install Nginx web server
sudo apt install -y nginx

# Start and enable Nginx
sudo systemctl start nginx
sudo systemctl enable nginx

# Check Nginx status
sudo systemctl status nginx

# Install UFW firewall
sudo apt install -y ufw

# Configure firewall rules
sudo ufw default deny incoming
sudo ufw default allow outgoing
sudo ufw allow from 192.168.30.0/24 to any port 22
sudo ufw allow 80/tcp
sudo ufw allow 443/tcp

# Enable firewall
sudo ufw enable

# Install Fail2Ban
sudo apt install -y fail2ban
sudo systemctl enable fail2ban
sudo systemctl start fail2ban

# Verify services
sudo ufw status
sudo systemctl status nginx
sudo fail2ban-client status
```

### Security Hardening Applied

- Root login disabled via SSH
- Password authentication disabled for SSH
- Key-based authentication only
- UFW firewall configured (default deny incoming)
- Fail2Ban intrusion prevention enabled
- Unnecessary services disabled
- Security headers configured in Nginx
- Automatic security updates enabled

---

## 10. Security Features

### Defense-in-Depth Implementation

| Security Layer | Implementation | Purpose |
|---------------|---------------|---------|
| Perimeter | pfSense Firewall | Control all traffic between networks |
| Network | VLAN Segmentation | Isolate different security zones |
| Host | UFW Firewall, Windows Firewall | Protect individual systems |
| Application | SSL/TLS, Hardening | Secure application layer |
| Data | File Permissions, Encryption | Protect stored data |
| Identity | Active Directory | Centralized authentication |

### Network Segmentation Benefits

- **Isolation:** DMZ servers cannot directly access LAN resources
- **Containment:** Compromise in one zone does not automatically affect others
- **Control:** Granular firewall rules between segments
- **Monitoring:** Separate logging and monitoring per segment

### Access Control Measures

- Strong password policies enforced via Group Policy
- Account lockout after failed attempts
- Key-based SSH authentication only
- Firewall rules limit network access by segment
- Administrator access restricted to management network

---

## 11. Security Testing

### Testing Methodology

The security testing was performed from the Kali Linux VM (KL-TEST01) located in the management network. The following tests were conducted to verify the security posture:

### Network Discovery Results

```bash
# Network scanning command
nmap -sn 192.168.10.0/24  # LAN scan
nmap -sn 192.168.20.0/24  # DMZ scan
nmap -sn 192.168.30.0/24  # MGMT scan
```

**Discovered Hosts:**
- LAN: 192.168.10.1 (pfSense), 192.168.10.10 (DC), 192.168.10.20 (Client)
- DMZ: 192.168.20.1 (pfSense), 192.168.20.10 (Web Server)
- MGMT: 192.168.30.1 (pfSense), 192.168.30.10 (Kali), 192.168.30.100 (ESXi)

### Firewall Rule Verification

| Test | Source | Destination | Expected | Result |
|------|--------|-------------|----------|--------|
| Test 1 | LAN | Internet | Allow | PASS |
| Test 2 | LAN | DMZ | Allow | PASS |
| Test 3 | LAN | MGMT | Block | PASS |
| Test 4 | DMZ | LAN | Block | PASS |
| Test 5 | DMZ | Internet | Allow | PASS |
| Test 6 | MGMT | All | Allow | PASS |

### Port Scanning Results

| Host | Open Ports | Service | Status |
|------|------------|---------|--------|
| DC-SRV01 | 53, 88, 389, 445 | DNS, Kerberos, LDAP, SMB | Expected |
| LN-SRV01 | 22, 80, 443 | SSH, HTTP, HTTPS | Expected |
| ESXi Host | 443, 902 | HTTPS, Management | Expected |

### Authentication Testing

- SSH password authentication disabled and verified
- Active Directory account lockout policy tested and working
- Firewall correctly blocks unauthorized access attempts

---

## 12. Monitoring and Logs

### Log Sources

| System | Log Type | Location | Purpose |
|--------|----------|----------|---------|
| Windows Server | Security Events | Event Viewer | Authentication tracking |
| Windows Server | System Events | Event Viewer | System status |
| pfSense | Firewall Logs | Web Interface | Traffic monitoring |
| pfSense | System Logs | Web Interface | Service status |
| Linux Server | auth.log | /var/log/auth.log | SSH authentication |
| Linux Server | syslog | /var/log/syslog | System messages |
| Linux Server | nginx logs | /var/log/nginx/ | Web access logs |

### Log Retention Policy

| Log Type | Retention | Rotation |
|----------|-----------|----------|
| Windows Security | 90 days | Weekly |
| Firewall Logs | 30 days | Daily |
| System Logs | 30 days | Weekly |
| Web Access Logs | 30 days | Daily |

### Monitoring Procedures

1. Daily review of authentication logs for failed attempts
2. Weekly review of firewall logs for blocked traffic patterns
3. Monthly review of security events and policy compliance
4. Real-time alerts for critical security events via Fail2Ban

---

## 13. Screenshots Section

### Network Diagram

```
[Insert Screenshot: Network Topology Diagram]
```

*Screenshot showing the complete network architecture with all VMs and connections.*

### pfSense Dashboard

```
[Insert Screenshot: pfSense Dashboard]
```

*Screenshot showing pfSense firewall dashboard with interface status and traffic graphs.*

### Active Directory Users and Computers

```
[Insert Screenshot: Active Directory Users]
```

*Screenshot showing the Active Directory Users and Computers console with organizational units.*

### Linux Server Terminal

```
[Insert Screenshot: Linux Server Terminal]
```

*Screenshot showing the Linux server terminal with system status commands output.*

### Windows Client Domain Join

```
[Insert Screenshot: Domain Join]
```

*Screenshot showing successful domain join of Windows client to Active Directory.*

### Firewall Rules

```
[Insert Screenshot: pfSense Firewall Rules]
```

*Screenshot showing the configured firewall rules in pfSense web interface.*

### Security Scan Results

```
[Insert Screenshot: Nmap Scan Results]
```

*Screenshot showing Nmap network scan results from Kali Linux.*

### Web Server Test

```
[Insert Screenshot: Nginx Test Page]
```

*Screenshot showing the Nginx default web page served from Linux server.*

---

## 14. Conclusion

This project successfully demonstrates the implementation of a secure virtual enterprise lab using VMware ESXi. The infrastructure includes proper network segmentation with VLANs, centralized authentication through Active Directory, a hardened Linux web server in a DMZ, and comprehensive security controls including firewall rules and monitoring.

The security testing confirmed that firewall rules are correctly blocking unauthorized traffic between network segments while allowing legitimate communications. The defense-in-depth approach provides multiple layers of security, ensuring that a compromise in one zone does not automatically lead to access in other zones.

This lab environment serves as an excellent platform for learning enterprise IT infrastructure management, network security, and system administration. It can be extended further with additional security tools, monitoring solutions, and automated deployment using infrastructure-as-code tools.

---

**End of Report**