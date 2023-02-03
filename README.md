# cover-letter-bot


## Introduction

This script generates a job application letter in PDF format, using the Python module `DocxTemplate` and `docxtopdf`. The user is prompted to provide inputs such as the focus of the letter, company details, hiring manager name, etc., which are then inserted into a predefined template to generate the final letter. The script also uses the `geopy` library to retrieve the company's address location data.

## Variables

The script defines a `Datax` class to hold the various variables used in the letter. The values of these variables are assigned based on the focus of the letter, which is input by the user.

## Template

The template is predefined, and the placeholders in the template are filled with the values of the variables from the `Datax` class.

## Conclusion

This script simplifies the process of generating a job application letter by automatically filling in the details, converting the letter to PDF format, and providing the option to retrieve location data for the company's address.
