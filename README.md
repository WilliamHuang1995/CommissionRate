[![Build Status](https://travis-ci.org/WilliamHuang1995/CommissionRate.svg?branch=master)](https://travis-ci.org/WilliamHuang1995/CommissionRate)
# Description
A small program that calculates the commission of each employee
  
# Program Logic
1. Display a GUI interface for user to select input file and output folder
2. The program parses the filepaths for input and output as specified by the user
3. There are a few specification of how the excel file should look.
    1. Employee name needs to be on the first column
    2. Customer name needs to be on the third column
    3. Product name needs to be on the fifth column
    4. Revenue needs to be on the tenth column  
4. Obviously for low coupling, these settings can be taken out and declared in a config file. But that's on the backlog for now. What matters is that it works right now.   
5. As there are different employees,customer, product, and revenues, in the python source folder, my data structure looks like this:

    ```
    employee:{
        customer:{
            product:revenue
        }
    }
    ```
    Below is an example of what the sheet gets parsed into after the process_data method 
    ``` 
    dictionary = {
        employee_one:{
            customer_one:{
                product_A: revenue,
                product_B: revenue,
                ...
            },
            customer_two:{
                product_A: revenue,
                ..
            }
        }
        employee_two:{
            ..
        }
    }
    ```
6. With the datastructure completed, it's much easier to parse the tree into a desired output onto excel. 

# Installation
If you haven't already downloaded Python 3.6, click [here](https://www.python.org/downloads/)  
run 
```
pip install openpyxl
```
  
then clone the repository 
```
git clone https://github.com/WilliamHuang1995/CommissionRate.git
```
navigate to python/src and run Commission.py and it should work.

# Backlog
* ~~create a java version~~
* ~~create a python version~~
* add more to the output.xlsx file
* add the variables into a config file
* separate the gui from the program logic
* create a javascript/c#/c/c++ version
* prettify the GUI interface

