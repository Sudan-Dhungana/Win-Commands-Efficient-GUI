import sys
import ctypes
import os
from PyQt6.QtWidgets import (QApplication, QMainWindow, QTabWidget, QPushButton, 
                            QVBoxLayout, QWidget, QScrollArea, QHBoxLayout, 
                            QLabel, QStyleFactory, QMessageBox, QTextEdit, QLineEdit)
from PyQt6.QtCore import Qt, QThreadPool, QRunnable, QTimer
from PyQt6.QtGui import QPalette, QColor
from functools import partial

class CommandRunner(QRunnable):
    """Worker class to run commands in a separate thread"""
    def __init__(self, command, is_admin, parent):
        super().__init__()
        self.command = command
        self.is_admin = is_admin
        self.parent = parent

    def run(self):
        try:
            if self.is_admin:
                ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", f"/c {self.command}", None, 1)
            else:
                os.system(self.command)
        except Exception as e:
            QApplication.postEvent(self.parent, Qt.QEvent(Qt.QEvent.Type.User))

class CommandCategory:
    """Class to organize command data"""
    def __init__(self):
        self.categories = {
            "System Tools": [
                ("msinfo32", "System Information", "Displays detailed configuration information about your computer's hardware and software.", False),
                ("taskmgr", "Task Manager", "View currently running processes, CPU usage, memory usage, and manage startup programs.", False),
                ("control", "Control Panel", "A central place to configure settings for your computer, including hardware, software, and user accounts.", False),
                ("sysdm.cpl", "System Properties", "Provides access to system settings like hardware, advanced, and computer name.", False),
                ("devmgmt.msc", "Device Manager", "Manage and update drivers for hardware devices connected to your computer.", True),
                ("regedit", "Registry Editor", "Allows you to view and modify the Windows registry, which contains configurations for the OS and applications.", True),
                ("services.msc", "Services", "Control services running on your system, start, stop, or configure them.", True),
                ("lusrmgr.msc", "Local Users and Groups", "Manage user accounts and groups on the local computer.", True),
                ("perfmon.msc", "Performance Monitor", "Monitor system performance in real-time, including CPU, memory, disk, and network usage.", False),
                ("gpedit.msc", "Group Policy Editor", "Edit group policies which control the working environment of user accounts and computer accounts.", True),
                ("diskmgmt.msc", "Disk Management", "Manage disk partitions, volumes, or dynamic disks; initialize new disks, and create or delete partitions.", True),
                ("compmgmt.msc", "Computer Management", "A suite of tools to manage local or remote computers, including Event Viewer, Performance, and Services.", True),
                ("cleanmgr", "Disk Cleanup", "Remove temporary files, system files, empty the Recycle Bin, and delete other unnecessary files to free up disk space.", False),
                ("ntmsmgr.msc", "Removable Storage", "Manage removable storage media and their libraries on your system.", True),
                ("powercfg.cpl", "Power Options", "Configure power settings to manage how your computer uses power, including sleep, hibernation, and display settings.", False),
                ("timedate.cpl", "Date and Time", "Set or change the date, time, and time zone for your computer.", False),
                ("mrt", "Malicious Software Removal Tool", "Scans your computer for specific, prevalent malicious software and helps remove infections.", False),
                ("mmc", "Microsoft Management Console", "A framework for system administrators to create and save sets of administrative tools for management tasks.", False),
                ("winver", "Check Windows Version", "Displays the version of Windows currently installed on your computer.", False),
                ("calc", "Calculator", "A simple calculator application for basic arithmetic operations.", False)
            ],
            "File Management": [
                ("explorer", "File Explorer", "The default file manager for Windows, allowing navigation through the file system.", False),
                ("notepad", "Notepad", "A simple text editor for basic text editing tasks.", False),
                ("write", "WordPad", "A more feature-rich text editor than Notepad, supporting RTF, Word documents, and more.", False),
                ("mspaint", "Paint", "A basic graphics painting program that comes with Windows.", False),
                ("charmap", "Character Map", "View and copy all the characters in a font, including special characters.", False),
                ("fonts", "Fonts Folder", "View, install, or uninstall fonts on your system.", False),
                ("winword", "Microsoft Word", "A word processing program for creating documents.", False),
                ("excel", "Microsoft Excel", "A spreadsheet program for data analysis and storage.", False),
                ("powerpnt", "Microsoft PowerPoint", "A presentation program used to create slideshows.", False)
            ],
            "Networking": [
                ("ipconfig", "IP Configuration", "Displays all current TCP/IP network configuration values and refreshes DHCP and DNS settings.", False),
                ("ping", "Ping Command", "Tests the reachability of a host on an IP network and measures the round-trip time for messages.", False),
                ("netstat", "Network Statistics", "Displays active TCP connections, routing tables, and interface statistics.", False),
                ("ncpa.cpl", "Network Connections", "Manage network adapters and connections on your computer.", False),
                ("dns", "DNS Management", "Manages DNS server settings and records.", False),
                ("firewall.cpl", "Windows Firewall", "Configure firewall settings to block or allow network traffic based on predefined rules.", False),
                ("cmd", "Command Prompt", "A command-line interpreter application available in most Windows operating systems.", False),
                ("powershell", "PowerShell", "A task automation and configuration management framework from Microsoft, consisting of a command-line shell and scripting language.", False)
            ],
            "Troubleshooting": [
                ("dxdiag", "DirectX Diagnostic Tool", "Checks your computer for DirectX compatibility and provides detailed information about your graphics hardware.", False),
                ("msconfig", "System Configuration", "Configure which programs run at startup, services, and boot options.", True),
                ("resmon", "Resource Monitor", "Shows detailed information about hardware and software resource use in real time.", False),
                ("taskschd.msc", "Task Scheduler", "Schedule automated tasks that perform actions at a specific time or when a certain event occurs.", True),
                ("sigverif", "File Signature Verification", "Verifies that system files have valid digital signatures, which can help identify modified or corrupted files.", True),
                ("sfc /scannow", "System File Checker", "Scans and repairs protected system files to ensure they are not corrupted or altered.", True),
                ("dism /online /cleanup-image /scanhealth", "DISM Scan Health", "Checks the Windows component store for corruption without making repairs.", True),
                ("dism /online /cleanup-image /restorehealth", "DISM Restore Health", "Repairs the Windows component store by downloading healthy files from Windows Update.", True),
                ("dism /online /cleanup-image /startcomponentcleanup", "DISM Component Cleanup", "Cleans up superseded components in the Windows component store to free disk space.", True),
                ("dism /online /cleanup-image /analyzecomponentstore", "DISM Analyze Store", "Analyzes the component store size and suggests cleanup actions.", True),
                ("dfrgui", "Defragment and Optimize Drives", "Defragments hard drives to improve file access times and optimizes SSDs.", True),
                ("msra", "Windows Remote Assistance", "Allows a user to receive or provide assistance over a network or internet connection.", False),
                ("mstsc", "Remote Desktop Connection", "Connects to another computer over a network or the internet to control it remotely.", False),
                ("eventvwr", "Event Viewer", "View and manage event logs which record significant events on your computer like errors, warnings, or information.", False)
            ],
            "Accessibility": [
                ("osk", "On-Screen Keyboard", "A virtual keyboard that appears on screen for touch or accessibility purposes.", False),
                ("magnify", "Magnifier", "A tool to zoom in parts of your screen for better visibility.", False),
                ("narrator", "Narrator", "A screen reader that reads text on the screen aloud for visually impaired users.", False),
                ("utilman", "Ease of Access Center", "Provides quick access to tools for making your computer easier to use for people with disabilities.", False)
            ]
        }

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.threadpool = QThreadPool()
        self.setWindowTitle("Windows Run Command Launcher")
        self.setGeometry(100, 100, 800, 600)
        self.commands = CommandCategory()
        self.all_commands = [(cmd, label, desc, admin) 
                           for cat in self.commands.categories.values() 
                           for cmd, label, desc, admin in cat]
        # Debounce timer for search
        self.search_timer = QTimer()
        self.search_timer.setSingleShot(True)
        self.search_timer.timeout.connect(self.perform_search)
        self.current_search_text = ""
        self.setup_ui()
        self.setup_styles()

    def setup_ui(self):
        """Initialize UI components"""
        main_layout = QVBoxLayout()
        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Search bar
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("Search commands...")
        self.search_bar.textChanged.connect(self.debounce_search)
        main_layout.addWidget(self.search_bar)

        # Tabs
        self.tabs = QTabWidget()
        for category, cmd_list in self.commands.categories.items():
            self.add_category_tab(category, cmd_list)
        
        # About Us tab
        about_tab = QWidget()
        about_layout = QVBoxLayout(about_tab)
        about_text = QTextEdit()
        about_text.setReadOnly(True)
        about_text.setText("""
            <h2>About Us</h2>
            <p>Welcome to the <b>Windows Run Command Launcher</b>, a tool designed to simplify system management for Windows users.</p>
            
            <h3>Collaboration</h3>
            <p>This application is the result of a collaboration between its creator and Grok, an AI assistant built by xAI. The creator envisioned a user-friendly launcher for essential Windows commands, while Grok provided optimization, category refinement, and feature enhancements.</p>
            
            <h3>Date</h3>
            <p>Created and updated as of <b>February 24, 2025</b>.</p>
            
            <h3>Purpose</h3>
            <p>Our goal is to empower users with quick access to system tools, file management utilities, networking commands, troubleshooting options, and accessibility features—all in one sleek interface.</p>
            
            <h3>Credits</h3>
            <p>- <b>Creator (Sudan Dhungana)</b>: Designed the initial concept and structure.<br>
               - <b>Grok (xAI)</b>: Assisted with code optimization, UI improvements, and added features like DISM commands and this About Us section.</p>
            
            <p>Thank you for using our tool!</p>
        """)
        about_layout.addWidget(about_text)
        self.tabs.addTab(about_tab, "About Us")
        main_layout.addWidget(self.tabs)

        # Search results
        self.search_results = QScrollArea()
        self.search_content = QWidget()
        self.search_layout = QVBoxLayout(self.search_content)
        self.search_results.setWidgetResizable(True)
        self.search_results.setWidget(self.search_content)
        main_layout.addWidget(self.search_results)
        self.search_results.hide()

    def setup_styles(self):
        """Configure application styles"""
        app.setStyle(QStyleFactory.create('Fusion'))
        
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor(45, 45, 45))
        palette.setColor(QPalette.ColorRole.WindowText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Base, QColor(35, 35, 35))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor(45, 45, 45))
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Button, QColor(55, 55, 55))
        palette.setColor(QPalette.ColorRole.ButtonText, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Highlight, QColor(42, 130, 218))
        palette.setColor(QPalette.ColorRole.HighlightedText, Qt.GlobalColor.black)
        QApplication.setPalette(palette)

    def add_category_tab(self, category, commands):
        """Create and populate a category tab"""
        tab = QWidget()
        scroll = QScrollArea()
        content = QWidget()
        
        layout = QVBoxLayout(tab)
        content_layout = QVBoxLayout(content)
        
        scroll.setWidgetResizable(True)
        scroll.setWidget(content)
        
        for command, label, desc, requires_admin in commands:
            row = self.create_command_row(command, label, desc, requires_admin)
            content_layout.addWidget(row)
        
        content_layout.addStretch()
        layout.addWidget(scroll)
        self.tabs.addTab(tab, category)

    def create_command_row(self, command, label, desc, requires_admin):
        """Create a single command row"""
        container = QWidget()
        layout = QHBoxLayout(container)
        
        button = QPushButton(label)
        button.setFixedWidth(200)
        button.clicked.connect(partial(self.execute_command, command, requires_admin))
        
        desc_label = QLabel(desc)
        desc_label.setWordWrap(True)
        
        layout.addWidget(button)
        layout.addWidget(desc_label)
        layout.addStretch()
        
        return container

    def execute_command(self, command, is_admin):
        """Execute command in a worker thread"""
        worker = CommandRunner(command, is_admin, self)
        self.threadpool.start(worker)

    def customEvent(self, event):
        """Handle custom events from thread"""
        if event.type() == Qt.QEvent.Type.User:
            QMessageBox.critical(self, "Error", "Failed to execute command. Check console for details.")

    def debounce_search(self, text):
        """Debounce search input to prevent rapid updates"""
        self.current_search_text = text
        self.search_timer.start(200)  # 200ms delay

    def perform_search(self):
        """Filter and display commands based on search input"""
        text = self.current_search_text
        # Clear previous results safely
        while self.search_layout.count():
            item = self.search_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        if text.strip():
            self.tabs.hide()
            self.search_results.show()
            filtered = [cmd for cmd in self.all_commands 
                       if text.lower() in cmd[1].lower() or text.lower() in cmd[2].lower()]
            for command, label, desc, requires_admin in filtered:
                row = self.create_command_row(command, label, desc, requires_admin)
                self.search_layout.addWidget(row)
            self.search_layout.addStretch()
        else:
            self.search_results.hide()
            self.tabs.show()

    def apply_styles(self):
        """Apply stylesheet to widgets"""
        self.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #4A82DA;
                background: #2B2B2B;
                border-radius: 5px;
            }
            QTabBar::tab {
                background: #353535;
                color: white;
                border: 1px solid #4A82DA;
                border-bottom: none;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                padding: 5px 15px;
                min-width: 100px;
            }
            QTabBar::tab:selected {
                background: #4A82DA;
                color: white;
            }
            QPushButton {
                background-color: #353535;
                color: white;
                border: 2px solid #4A82DA;
                border-radius: 10px;
                padding: 5px;
                min-height: 30px;
            }
            QPushButton:hover {
                background-color: #4A82DA;
            }
            QLabel {
                color: white;
            }
            QTextEdit {
                background-color: #353535;
                color: white;
                border: 1px solid #4A82DA;
                border-radius: 5px;
                padding: 5px;
            }
            QLineEdit {
                background-color: #353535;
                color: white;
                border: 1px solid #4A82DA;
                border-radius: 5px;
                padding: 5px;
            }
        """)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.apply_styles()
    window.show()
    sys.exit(app.exec())