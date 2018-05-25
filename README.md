# app-stopwatch
App for debugging VBA performance

Contains a .bas file with code for starting, stopping and printing time. There are two constants which regulates where the logging is made.
* PRINT_DEBUG = logging to the immediate window.
* PRINT_LOG = logging to the infolog table.

Usage:

Call StopWatch.Start("Class_Initialize.Setup") 'Name of timer which also is printed out
Call StopWatch.PrintTime("Class_Initialize.Setup") 'Timer is removed after print