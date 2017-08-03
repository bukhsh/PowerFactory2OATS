# PowerFactory2OATS
A collection of Python scripts that can convert a network data from DIgSILENT PowerFactory into a spreadsheet (a readable data format for optimisation and analysis toolbox for power systems (OATS): a software developed at The University of Strathclyde). The converted data can also be used to produce a Matpower file, however, a script to produce a Matpower file is not provided here. 

Note that access to DIgSILENT PowerFactory is required to perform this conversion.

## Using this filter
The top level script is 'run_file_PF.py'. This script needs to run inside DIgSILENT PowerFactory. Number of 'csv' files will be produced after this script is successfully executed. The csv files containts information about the network e.g. chracteristics of branches and generators.

The script 'run_file_writexlsx.py' can be used to convert the csv files into a spreadsheet. 


## Disclaimer

Contents of this archieve are for research and educational purposes. I have made every effort to ensure its accuracy and correctness. However the material in this archieve is distributed with no gurantees regarding correctness, its quality or fitness for a specific application.

## Author

* **Waqquas Bukhsh** 


## License

This project is licensed under the GNU General Public License v3.0 - see the LICENSE file for details.

## Acknowledgments

* Arnaud Vergnol, ELIA.

