<a name="readme-top"></a>

[![Contributors][contributors-shield]](https://github.com/gelndjj/Active_Directory_Automate/graphs/contributors)
[![Forks][forks-shield]](https://github.com/gelndjj/Active_Directory_Automate/forks)
[![Stargazers][stars-shield]](https://github.com/gelndjj/Active_Directory_Automate/stargazers)
[![Issues][issues-shield]](https://github.com/gelndjj/Active_Directory_Automate/issues)
[![MIT License][license-shield]](https://github.com/gelndjj/Active_Directory_Automate/blob/main/LICENSE)
[![LinkedIn][linkedin-shield]](https://www.linkedin.com/in/jonathanduthil/)


<!-- PROJECT LOGO -->
<br />
<div align="center">
  <a href="https://github.com/gelndjj/Active_Directory_Automate">
    <img src="https://github.com/gelndjj/Active_Directory_Automate/blob/main/resources/image0.png" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">Active Directory Automate (Windows Server Machine)</h3>

  <p align="center">
    A Python application to handle Active Directory tasks
    <br />
    <a href="https://github.com/gelndjj/Active_Directory_Automate"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    ·
    <a href="https://github.com/gelndjj/Active_Directory_Automate/issues">Report Bug</a>
    ·
    <a href="https://github.com/gelndjj/Active_Directory_Automate/issues">Request Feature</a>
  </p>
</div>



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#contributing">Contributing</a></li>
    <li><a href="#license">License</a></li>
    <li><a href="#contact">Contact</a></li>

  </ol>
</details>


<!-- ABOUT THE PROJECT -->
## About The Project
<div align="center">
<img src="https://github.com/gelndjj/Active_Directory_Automate/blob/main/resources/image0.png" alt="Logo" width="128" height="128">
</br>
Active Directory Automate is a sophisticated application designed to streamline the process of user management within Active Directory (AD). Born out of the necessity to eliminate the repetitive and manual task of creating AD users, this tool stands as a testament to the power of automation in the IT administration field.</br> 
</br>
<img src="https://github.com/gelndjj/Active_Directory_Automate/blob/main/resources/main_ada.png" alt="Screenshot" width="927" height="852">
</br>
</br>
At the heart of Active Directory Automate is its ability to create and manage SQL databases that serve as the repository for user information. With functionality that enables you to add, edit, and modify user details, the application seamlessly synchronizes this data with Active Directory, ensuring a synchronized and up-to-date user management system.</br>
</br>
The principal aim of this project is to significantly reduce the administrative overhead associated with managing thousands of users in Active Directory</br>
</br>
<img src="https://github.com/gelndjj/Active_Directory_Automate/blob/main/resources/main_ada_2.png" alt="Screenshot" width="927" height="852">
</br>
</div>

<p align="right">(<a href="#readme-top">back to top</a>)</p>

<!-- USAGE EXAMPLES -->

# Usage 

1. Launching the Application:
+ Critical Note: The executable (EXE) file must be launched directly on the Windows Server machine where Active Directory services are hosted. This is essential for the tool to function correctly and securely interact with Active Directory.
+ Do not attempt to run the executable from client machines. Doing so may result in improper functionality and security risks.
2. Performing Tasks:
+ After launching the application on the server, follow the on-screen instructions to perform various Active Directory management tasks.
+ Security and Compliance:
Ensure that you have the necessary permissions and comply with your organization's IT policies when using this tool on the server.

# Usage Documentation

## Table of Contents
- [Record Management](#record-management)
- [PowerShell Commands](#powershell-commands)
- [Python Commands](#python-commands)
- [Database Storing](#database-storing)
- [Organisation Unit Management](#organisation-unit-management)
- [Auto Filling](#auto-filling)

## Record Management

### Add Record
- **Description**: Creates a new user record in the application with the details provided in the input fields.
- **Usage**: Complete the input fields under 'Record' and click this button to save a new record.

### Update Record
- **Description**: Modifies an existing user record.
- **Usage**: After editing details in the 'Record' section, click this button to apply the updates.

### Clear Entry Boxes
- **Description**: Clears all the input fields in the 'Record' section.
- **Usage**: Click this to reset the form and prepare for new input.

### Reassign IDs
- **Description**: Automatically reassigns IDs to all records in the database to maintain a sequential order.
- **Usage**: Use this button to reorganize the ID sequence after record deletions.

### Delete Duplicates
- **Description**: Scans the database for duplicate records and removes them.
- **Usage**: Click this button to clean up the database by removing duplicates.

### Remove Record
- **Description**: Deletes the selected user record from the database.
- **Usage**: Select a record and use this button to permanently delete it from the database.

### Remove All Records
- **Description**: Deletes all records from the database.
- **Usage**: Use this with caution; it will completely empty the database of user records.

## PowerShell Commands

### Sync Database to AD
- **Description**: Synchronizes the current database entries with Active Directory. The selected databsse in the DataBase Storing section will be synchronized. 
- **Usage**: Click this to update Active Directory with the information from the application database.

### Edit Sync DB Script
- **Description**: Opens the script used for database synchronization for editing.
- **Usage**: Click this if you need to modify the PowerShell script parameters or behavior.

### Change User Password
- **Description**: Initiates a prompt to change the password for a single or multiple selected users from the database then synchronize the changes to Active Directory.
- **Usage**: Select one or many user(s) and click this to begin the password change process.

### Enable/Disable User(s)
- **Description**: Makes the account of the selected user(s) Enable if Disable and Disable if Enable.
- **Usage**: Select one or many user(s) and click this to begin the enable/disable account process.

## Python Commands

### Add One Fake Record
- **Description**: Inserts a placeholder record with generated information into the database.
- **Usage**: Click this to add a single dummy record for testing or demonstration purposes.

### Add Multiple Fake Records
- **Description**: Generates and adds multiple placeholder records to the database.
- **Usage**: Click this to insert a batch of dummy records.

### Export Database to CSV
- **Description**: Exports the entire database to a CSV file named 'database_name'-YYYYMMDD-HH-MM-SS.
- **Usage**: Use this button to create a CSV backup or for data analysis purposes.

### Import Record from CSV
- **Description**: Imports records into the database from a CSV file.
- **Usage**: Click this and select a CSV file to upload and add records from it to the database.

## Database Storing

### Create New Database
- **Description**: Initializes a new, empty database for storing user records.
- **Usage**: Click this to set up a fresh database, typically used when starting from scratch.
- **IMPORTANT**: Place the database at the root of the application.

### Browse Database
- **Description**: Opens a file dialog to explore and manage the database files.
- **Usage**: Use this to locate and open an existing database file.
- **Detail**: Copies the selected database at the root of the application.

### Backup Current Database
- **Description**: Creates a backup copy of the current database and name it 'database_name'-YYYYMMDD-HH-MM-SS.
- **Usage**: Click this to ensure you have a recent backup of your user data.

## Organisation Unit Management

### Select Organisation Unit
- **Description**: Allows the selection of an organizational unit from Active Directory to manage.
- **Usage**: Click to display and select from available organizational units.

### Create Organisation Unit Path
- **Description**: Initiates a prompt to set up an existing organizational unit path in order to synchronize with Active Directory.
- **Usage**: Click this after specifying the details to create a new organizational path.

### Scan AD OU's Users
- **Description**: Scans Active Directory to find the Organizational Units (OUs) where users are already present.
- **Usage**: Click this button to perform a search in Active Directory and retrieve a list of OUs that contain user accounts. This can be useful for auditing purposes or to ensure users are correctly organized within AD.
- **Expected Outcome**: After scanning, the application will display the OUs along with associated user details from Active Directory.

## Auto Filling

### Create Address
- **Description**: Automatically generates and fills in an address for the selected record.
- **Usage**: Select a user record and click this to populate the address fields.

### Create SAM@domain
- **Description**: Constructs and fills in the SAM account name for a user based on the domain.
- **Usage**: With a user record selected, click this to auto-generate the SAM account name.


<!-- GETTING STARTED -->
## Standalone APP

Install pyintaller
```
pip install pyinstaller
```
Generate the standalone app
```
pyinstaller --onefile your_script_name.py
```


<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

If you have a suggestion that would make this better, please fork the repo and create a pull request. You can also simply open an issue with the tag "enhancement".


1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

<p align="right">(<a href="#readme-top">back to top</a>)</p>

## Built With

<a href="https://www.python.org">
<img src="https://github.com/gelndjj/Active_Directory_Automate/blob/main/resources/py_icon.png" alt="Icon" width="32" height="32">
</a>
&nbsp
<a href="https://customtkinter.tomschimansky.com">
<img src="https://github.com/gelndjj/Active_Directory_Automate/blob/main/resources/ctk_icon.png" alt="Icon" width="32" height="32">
</a>
&nbsp
<a href="https://learn.microsoft.com/en-us/powershell/scripting/learn/ps101/01-getting-started?view=powershell-7.3">
<img src="https://github.com/gelndjj/Active_Directory_Automate/blob/main/resources/ps_icon.png" alt="Icon" width="32" height="32">
</a>
<p align="right">(<a href="#readme-top">back to top</a>)</p>
    

<!-- LICENSE -->
## License

Distributed under the GNU GENERAL PUBLIC LICENSE. See `LICENSE.txt` for more information.

<p align="right">(<a href="#readme-top">back to top</a>)</p>



<!-- CONTACT -->
## Contact


[LinkedIn](https://www.linkedin.com/in/jonathanduthil/)

<p align="right">(<a href="#readme-top">back to top</a>)</p>


<!-- MARKDOWN LINKS & IMAGES -->
<!-- https://www.markdownguide.org/basic-syntax/#reference-style-links -->
[contributors-shield]: https://img.shields.io/github/contributors/othneildrew/Best-README-Template.svg?style=for-the-badge
[contributors-url]: https://github.com/othneildrew/Best-README-Template/graphs/contributors
[forks-shield]: https://img.shields.io/github/forks/othneildrew/Best-README-Template.svg?style=for-the-badge
[forks-url]: https://github.com/othneildrew/Best-README-Template/network/members
[stars-shield]: https://img.shields.io/github/stars/othneildrew/Best-README-Template.svg?style=for-the-badge
[stars-url]: https://github.com/othneildrew/Best-README-Template/stargazers
[issues-shield]: https://img.shields.io/github/issues/othneildrew/Best-README-Template.svg?style=for-the-badge
[issues-url]: https://github.com/othneildrew/Best-README-Template/issues
[license-shield]: https://img.shields.io/github/license/othneildrew/Best-README-Template.svg?style=for-the-badge
[license-url]: https://github.com/othneildrew/Best-README-Template/blob/master/LICENSE.txt
[linkedin-shield]: https://img.shields.io/badge/-LinkedIn-black.svg?style=for-the-badge&logo=linkedin&colorB=555
[linkedin-url]: https://linkedin.com/in/othneildrew
[product-screenshot]: images/screenshot.png
[Next.js]: https://img.shields.io/badge/next.js-000000?style=for-the-badge&logo=nextdotjs&logoColor=white
[Next-url]: https://nextjs.org/
[React.js]: https://img.shields.io/badge/React-20232A?style=for-the-badge&logo=react&logoColor=61DAFB
[React-url]: https://reactjs.org/
[Vue.js]: https://img.shields.io/badge/Vue.js-35495E?style=for-the-badge&logo=vuedotjs&logoColor=4FC08D
[Vue-url]: https://vuejs.org/
[Angular.io]: https://img.shields.io/badge/Angular-DD0031?style=for-the-badge&logo=angular&logoColor=white
[Angular-url]: https://angular.io/
[Svelte.dev]: https://img.shields.io/badge/Svelte-4A4A55?style=for-the-badge&logo=svelte&logoColor=FF3E00
[Svelte-url]: https://svelte.dev/
[Laravel.com]: https://img.shields.io/badge/Laravel-FF2D20?style=for-the-badge&logo=laravel&logoColor=white
[Laravel-url]: https://laravel.com
[Bootstrap.com]: https://img.shields.io/badge/Bootstrap-563D7C?style=for-the-badge&logo=bootstrap&logoColor=white
[Bootstrap-url]: https://getbootstrap.com
[JQuery.com]: https://img.shields.io/badge/jQuery-0769AD?style=for-the-badge&logo=jquery&logoColor=white
[JQuery-url]: https://jquery.com 
