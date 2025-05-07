# Project Title: SharePoint Data Submission Application

## Description
This Flask application allows users to log in with their SharePoint credentials, submit data through a form, and view the submitted data in a table format. The application interacts with a SharePoint list to store and retrieve data.

## Project Structure
```
Site
├── templates
│   ├── form.html        # HTML structure for the data submission form
│   ├── login.html       # HTML structure for the login page
│   └── main.html        # HTML structure for displaying the SharePoint list
├── app.py               # Main application file
└── README.md            # Documentation for the project
```

## Setup Instructions
1. **Clone the repository**:
   ```
   git clone <repository-url>
   cd Site
   ```

2. **Install the required packages**:
   Ensure you have Python and pip installed, then run:
   ```
   pip install Flask SharePlum
   ```

3. **Configure the application**:
   Update the `app.py` file with your SharePoint site URL and credentials.

4. **Run the application**:
   ```
   python app.py
   ```
   The application will start on `http://127.0.0.1:5000/`.

## Usage Guidelines
- Navigate to the login page to enter your SharePoint credentials.
- Upon successful login, you will be redirected to the main page where you can view the SharePoint list.
- Use the button at the top of the main page to navigate to the form for submitting new data.
- After submitting data, you will be redirected back to the main page to see the updated list.

## Notes
- Ensure that your SharePoint account has the necessary permissions to access and modify the list.
- This application is intended for educational purposes and may require further enhancements for production use.