@echo off
python setup.py clean --all
python setup.py sdist bdist_wheel
