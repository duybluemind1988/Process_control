#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Dec 10 17:40:19 2020

@author: ad
"""

import matplotlib.pyplot as plt
import numpy as np

def plot_line1():
    print("plot sin line")
    x = np.linspace(0, 20, 100)  # Create a list of evenly-spaced numbers over the range
    plt.plot(x, np.sin(x))       # Plot the sine of each x point
    plt.pause(0.3) 
    #plt.show()                   # Display the plot
    plt.close()
plot_line1()

def hello_world(firstname,lastname):
    var=12
    print("my name is {}, {}".format(firstname,lastname))
    var=24
    sum1=var + 250
    print(sum1)
    
hello_world('Krish','Naik')
#!pip install streamlit

def calculation(x,y):
    print("sum: ",x+y)

calculation(2,3)

def plot_line2():
    print("plot line")
    plt.plot([1, 2, 3, 4])
    plt.ylabel('some numbers')
    plt.pause(0.3) 
    #plt.show()
    plt.close()
plot_line2()