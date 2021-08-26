from setuptools import setup

setup(
    name='uppraisal',
    version='0.0.1',
    url='https://github.com/cisglee/uppraisal',
    packages=['tools'],
    license='MIT',
    author='CISG Lee',
    author_email='clee@rsm.nl',
    description='Bulk upload assignment comments and grades to Canvas LMS',
    classifiers=[
        'License :: OSI Approved :: MIT License',
        'Natural Language :: English',
        'Operating System :: OS Independent',
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3.6',
        'Programming Language :: Python :: 3.7',
        'Programming Language :: Python :: 3.8',
        'Programming Language :: Python :: 3.9',
    ],
    install_requires=['requests', 'xlrd>=1.0.0'],
)
