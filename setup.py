from setuptools import setup, find_packages  

setup(  
    name = "pyAddin", 
    version = "0.1",  
    keywords = ("python", "vba", "excel", "addin"),  
    description = "python-vba Excel addin",  
    long_description = "Excel addin template combined VBA with Python",
    license = "MIT Licence", 
    author = "dothinking",  
    author_email = "train8808@gmail.com",  
    packages = find_packages(exclude=["test", "dist"]),  
    include_package_data = True, 
    install_requires=[
        'argparse',
        'PyYAML>=3.13',
        'pywin32>=224'
    ],
    zip_safe=False,
    platforms = "windows",  
    entry_points = {  
        'console_scripts': [  
            'pyaddin = pyaddin.main:main'  
        ]  
    }  
)