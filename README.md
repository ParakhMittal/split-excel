# split-excel
This repository contains python code to split data in an excel workbook (for all worksheets) into multiple excel files based on column value.

<!-- TABLE OF CONTENTS -->
<details open="open">
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
      </ul>
    </li>
    <li><a href="#contributing">Contributing</a></li>
  </ol>
</details>

<!-- ABOUT THE PROJECT -->
## About The Project

This project contains the utility to split an excel workbook (.xlsx) into multiple excel workbooks based on the column value. The source excel file can have multiple worksheets. All worksheets must have first row for column headers. 

The utility looks for unique values of the filter column across all worksheets and then create new workbook (e.g. Split-{Value1}.xlsx) for each unique value such that the newly generated excel workbook has records only for filtered value (e.g. Value1).
The utility also copies the cell format from source excel workbook to all destination workbook.  

### Built With
This utility is build using
* [Python](https://www.python.org/downloads/)



<!-- GETTING STARTED -->
## Getting Started

To get a local copy up and running follow these simple example steps.

### Prerequisites

In order to run the utility on your local, you need to follow below-mentioned steps.

*  <b>Python</b>

    Download and install python from [here](https://www.python.org/downloads/)
    
### Installation

1. Clone the repo
   ```sh
   git clone https://github.com/ParakhMittal/split-excel.git
   ```
2. Set the location of installation directory (e.g. C:\Program Files\Python38) of Python in the environment variable PATH.
3. Set the location of installation directory (e.g. C:\Program Files\Python38\Scripts) of Python in the environment variable PATH.
4. Install pipenv 

        pip install pipenv

5. Create the virtual environment 

        pipenv --three
        
6. Install dependencies

        pipenv install
 
7. Activate the virtual environment (created by pipenv)

        pipenv shell

8. Run the utility

        python -m SplitExcel.py <src_file> <filter_col_name> <des_directory>

9. Exit the virtual environment.

        exit
        
<!-- CONTRIBUTING -->
## Contributing

Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/my-feature-branch`)
3. Commit your Changes (`git commit -m 'Add some feature'`)
4. Push to the Branch (`git push origin feature/my-feature-branch`)
5. Open a Pull Request





