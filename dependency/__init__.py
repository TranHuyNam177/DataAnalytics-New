CompanyName = 'Công Ty Cổ Phần Chứng Khoán Phú Hưng'
CompanyAddress = 'Tầng 21, Phú Mỹ Hưng Tower, 08 Hoàng Văn Thái, phường Tân Phú, quận 7, Thành phố Hồ Chí Minh'
CompanyPhoneNumber = 'Điện thoại : (+84 28)  5413 5479  |  Fax: (+84 28) 5413 5472'
CompanyEmail = 'Email: info@phs.vn / support@phs.vn  - Web: www.phs.vn'
CompanyCode = '022'

# IMPORT PACKAGES

import numpy as np
import pandas as pd

pd.set_option('display.max_rows',None,'display.max_columns',None,'display.width',None)

import cv2
import numexpr as ne
import openpyxl
import os
import sys
from os import listdir
from os.path import isfile,isdir,join,realpath,dirname
import win32com as win32
from win32com.client import Dispatch
import time
import datetime as dt
import requests
import json
import holidays
import calendar
from typing import Union,Callable
import itertools
import pyodbc
import xlwings as xw
import xlsxwriter
import docx
import openpyxl
import unidecode
import shutil
from PIL import Image
import csv
import re
import unidecode
import zeep
from zeep import Client
import threading
import traceback

###############################################################################

import matplotlib as mpl

mpl.use('Agg')
import matplotlib.pyplot as plt

plt.ioff()
from matplotlib.ticker import MaxNLocator,FuncFormatter,ScalarFormatter,FormatStrFormatter
import matplotlib.ticker
from matplotlib.dates import DateFormatter
import matplotlib.transforms as transforms
from pylab import yticks

import seaborn as sns

###############################################################################

import scipy as sc
from scipy import stats
from scipy.stats import rankdata
from scipy.interpolate import interp1d

###############################################################################

from sklearn.model_selection import train_test_split
from sklearn.linear_model import LinearRegression
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA

###############################################################################

import pytesseract
pytesseract.pytesseract.tesseract_cmd = join(dirname(__file__),'tesseract','tesseract')

###############################################################################

import base64
import struct
import calendar

###############################################################################

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchFrameException
from selenium.common.exceptions import NoSuchWindowException
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait,Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options

import scrapy
from bs4 import BeautifulSoup
from urllib.request import urlopen
import lxml

###############################################################################

import multiprocessing

###############################################################################

import tkinter as tk
