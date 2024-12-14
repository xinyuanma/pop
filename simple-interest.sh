#!/bin/bash

# Script to calculate simple interest

# Prompt the user to enter the principal amount
echo "Enter the principal amount (P):"
read principal

# Prompt the user to enter the annual interest rate (in percentage)
echo "Enter the annual interest rate (R):"
read rate

# Prompt the user to enter the time in years (T)
echo "Enter the time in years (T):"
read time

# Calculate simple interest using the formula: SI = (P * R * T) / 100
simple_interest=$(echo "scale=2; ($principal * $rate * $time) / 100" | bc)

# Display the result
echo "The Simple Interest is: $simple_interest"
