# PowerFactory2OATS
A collection of Python scripts that can convert a network data from DIgSILENT PowerFactory into a spreadsheet (a readable data format for optimisation and analysis toolbox for power systems (OATS): a software developed at The University of Strathclyde). Note that a access to DIgSILENT PowerFactory is required to perform this conversion.

## Using this filter
The top level script is 'run_file_PF.py'. This script needs to run inside DIgSILENT PowerFactory. Number of 'csv' files will be produced after this script is successfully executed. The csv files containts information about the network e.g. chracteristics of branches and generators.

The script 'run_file_writexlsx.py' can be used to convert the csv files into a spreadsheet. 


## Disclaimer

Contents of this archieve are for research and educational purposes. I have made every effort to ensure its accuracy and correctness. However the material in this archieve is distributed with no gurantees regarding correctness, its quality or fitness for a specific application.


## References
[1] R. D. Zimmerman, C. E. Murillo-Sanchez, and R. J. Thomas, "MATPOWER: Steady-State Operations, Planning and Analysis Tools for Power Systems Research and Education," Power Systems, IEEE Transactions on, vol. 26, no. 1, pp. 12–19, Feb. 2011. 
[2] C. M. Ferreira, et al., "Transient stability preventive control of an electric power system using a hybrid method,", 12th International Middle East Power System Conference 2008.

## Author

* **Waqquas Bukhsh** 


## License

This project is licensed under the GNU General Public License v3.0 - see the [LICENSE.md](LICENSE.md) file for details

## Acknowledgments

* Arnaud Vergnol, ELIA.

