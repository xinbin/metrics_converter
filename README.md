# metrics_converter
A converting tool for some metrics stuff!

## How to install/set up
1. set up a virtual environment. if you use anaconda, run `conda create -n metrics_converter pip`   
This will install pip, which you will use to install any required libs.
2. If you did not create an anaconda env, skip this step. activate the newly created env with `source activate metrics_converter`
3. Install the required `openpyxl` lib, run `pip install openpyxl`

## How to run/use

To view help menu:
```$ python convert.py -h```


Run command with full path to file. Output file will be created in the working directory:
```$ python convert.py <full path to file>```

