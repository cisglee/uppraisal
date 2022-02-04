from setuptools import setup

setup(
    name='uppraisal',
    version='0.0.1',
    url='https://github.com/cisglee/uppraisal',
    packages=['uppraisal'],
    license='MIT',
    author='CISG Lee',
    author_email='clee@rsm.nl',
    description='Bulk upload assignment comments and grades to Canvas LMS',
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
    ],
    python_requires='~=3.6',
    install_requires=['requests', 'urllib3', 'tqdm', 'openpyxl', 'et-xmlfile', 'bs4'],
)
