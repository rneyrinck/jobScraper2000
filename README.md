# Application Viewer

## Overview

**Application Viewer** is a Python-based GUI application designed to streamline your job application process. It helps you manage and tailor your resumes and cover letters for each job posting retrieved using the CareerJet API. Additionally, it stores your personal information for quick copy/paste when filling out job applications.

## Features

- **Job Search Integration**: Search for jobs using keywords and location via the CareerJet API.
- **Tailored Resumes and Cover Letters**: Automatically generate customized resumes and cover letters for each job posting.
- **Application Tracking**: Keep track of your job applications with status updates.
- **Information Management**: Store and easily access your skills, education, websites, and social profiles.
- **Drag-and-Drop Documents**: Quickly attach tailored resumes and cover letters to your applications.
- **Preview Documents**: View your tailored resumes and cover letters before sending.
- **Copy Personal Information**: Easily copy skills, education, and other details to your clipboard.

## Installation

### Prerequisites

- Python 3.x
- [pip](https://pip.pypa.io/en/stable/installation/) package manager

### Dependencies

The application requires the following Python packages:

- `pandas`
- `requests`
- `spacy`
- `beautifulsoup4`
- `python-docx`
- `PyQt5`

### Installation Steps

1. **Clone the Repository**

   ```bash
   git clone https://github.com/yourusername/application-viewer.git
   cd application-viewer
   ```

2. **Create a Virtual Environment (Optional but Recommended)**

   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install Required Packages**

   ```bash
   pip install -r requirements.txt
   ```

   If you don't have a `requirements.txt`, you can install the packages individually:

   ```bash
   pip install pandas requests spacy beautifulsoup4 python-docx PyQt5
   ```

4. **Download spaCy Language Model**

   The application uses spaCy's English language model. Download it using:

   ```bash
   python -m spacy download en_core_web_sm
   ```
   
## Usage

1. **Run the Application**

   ```bash
   python application_viewer.py
   ```

2. **First-Time Setup**

   - Upon first run, you will be greeted with a welcome message and instructions.
   - Upload your base resume template in `.docx` format via `File > Upload Resume`.
   - The application will save your resume template path for future use.

3. **Generating Applications**

   - Navigate to `File > Generate Applications` to start fetching job listings.
   - Enter your search parameters, including keywords, location, start page, end page, and page size.
   - The application will retrieve job postings and generate tailored resumes and cover letters for each.

4. **Reviewing Applications**

   - The main table displays all your applications.
   - Click on an application to view details, open the job link, or access tailored documents.
   - Update the application status using the dropdown menu.

5. **Managing Documents**

   - Use the **Drag Resume** and **Drag Cover Letter** buttons to drag and drop your documents into applications or emails.
   - Preview your tailored resume and cover letter using the **Preview** buttons.

6. **Copying Personal Information**

   - Quickly copy your job-specific skills, base skills, combined skills, education, websites, and social profiles using the provided buttons.

## Configuration

The application saves configuration settings in `config.json`, including your resume template path and first-run status.

## File Structure

- `application_viewer.py`: Main application script.
- `config.json`: Configuration file for storing settings.
- `Applications.csv`: CSV file storing your job applications data.
- `tailored_documents/`: Directory where tailored resumes and cover letters are saved.
- `Cover_Letter_Template.txt`: Text file template for generating cover letters.

## Templates

### Resume Template

- Upload your base resume in `.docx` format.
- Ensure it contains a **Skills** section that can be updated with job-specific skills.

### Cover Letter Template

- The `Cover_Letter_Template.txt` file is used to generate tailored cover letters.
- It should contain placeholders that will be replaced with actual data.

**Example Placeholders:**

- `{your_name}`
- `{your_address}`
- `{your_city_state_zip}`
- `{your_email}`
- `{your_phone}`
- `{date}`
- `{hiring_manager_name}`
- `{company_name}`
- `{company_address}`
- `{company_city_state_zip}`
- `{job_title}`
- `{skills_sentence}`

## Applicant Information

Your personal information is stored in the `applicant_info` dictionary within the code. Update it with your details:

```python
applicant_info = {
    'name': 'Your Name',
    'address': 'Your Address',
    'city_state_zip': 'Your City, State ZIP',
    'email': 'your.email@example.com',
    'phone': 'Your Phone Number',
    'websites': [
        'https://yourwebsite.com',
        'https://yourportfolio.com'
    ],
    'education': 'Your Education Details',
    'social_profiles': [
        'LinkedIn: https://linkedin.com/in/yourprofile',
        'GitHub: https://github.com/yourusername'
    ],
}
```

## Customization

- **Base Skills**: Update the `base_skills` list with your own skills to improve keyword matching.
- **Status Options**: The application uses predefined status options for tracking applications. Adjust the `status_options` list if needed.


## Dependencies

Ensure all dependencies are installed:

```bash
pip install pandas requests spacy beautifulsoup4 python-docx PyQt5
```

## Contributing

Contributions are welcome! If you wish to contribute to the project, feel free to submit issues or pull requests on GitHub.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

- **Robert Neyrinck**
  - Email: [robert.a.neyrinck@gmail.com](mailto:robert.a.neyrinck@gmail.com)
  - LinkedIn: [https://linkedin.com/in/robert-neyrinck](https://linkedin.com/in/robert-neyrinck)
  - GitHub: [https://github.com/rneyrinck](https://github.com/rneyrinck)

---

Feel free to reach out if you have any questions or need assistance with the application.
