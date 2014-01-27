This file contains help material on running the Aspose Java Examples Dashboard application.
Please read the portion specific to your Operating System.

Windows XP, Vista, 7
=========================================================

How to Run the application:
---------------------------------------------------------
- Double click on Start.bat. It should run the application.
- See "Troubleshooting" section, if it gives any error.
- See "How to Use the application (All Platforms)" section for more details.

System Requirements:
---------------------------------------------------------
- Sun JRE 1.7

Troubleshooting:
---------------------------------------------------------
- java command not found
Please install the latest version of Sun Java (JRE).
It is recommended to use the installer (.exe) and choose the default location.
Run the "Start.bat" again, it should run the Examples Dashboard.

- Java 1.7 not found
Please edit the "Start.bat" file. In ":manual" section, set the path manually

Linux (Ubuntu/Fedora)
=========================================================

How to Run the application:
---------------------------------------------------------
- Edit "start.sh" and set the value of SUN_JRE_1_7_HOME variable to Sun Java (JRE) 1.7 home folder
- Open terminal, change the directory to where you extracted the files
- Run the following command
- sh start.sh
- It should run the application.
- See "Troubleshooting" section, if it gives any error.
- See "How to Use the application (All Platforms)" section for more details.

Troubleshooting:
---------------------------------------------------------
- Sun Java 1.7 is not installed.
	- Download the zip archive of Sun JRE 1.7
	- Extract contents of the archive to /usr/lib/jvm/ folder or any folder of your choice
	- Update the variable SUN_JRE_1_7_HOME  in "start.sh" e.g. "/usr/lib/jvm/jre1.7.0_15"
	- It will not interfere with default Open JDK installation, because the bash script uses
	  the java command from the path you specified explicitly.

- Input, Output files and Browse Source giving error message
	- It uses Java's Desktop API to open files and folders with default OS registered program.
	- The Java Desktop API is dependent on libgnome package. Please install it to fix this issue.
	- Install command for Ubuntu: sudo apt-get install libgnome2-0
	- Install command for Fedora: sudo yum install libgnome
	
	- If there is no program installed to open the specific file, it will still give error message.
	  For example, if you try to open MSG or EML file in Ubuntu/Fedora, the error message will still appear.
	  For file types like DOC, DOCX, PPT, XLS etc, Ubuntu/Fedora may open these in Libre/Open Office

How to Use the application (All Platforms)
=========================================================
- On initial run of batch/script file, it will download the required files from Github 
  and will automatically run the Examples Dashboard application.
- Select any product from left tree and press "Update Examples" button.
  It will download the source code of selected products and the required libraries.
  On first download, it might take few minutes.
- You can update the examples anytime, to get the updates and newly added examples.