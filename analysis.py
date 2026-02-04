import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
import numpy as np
from scipy.optimize import curve_fit
from config import *
import datetime

curtime = datetime.datetime.now(datetime.UTC).isoformat()
curyear = curtime[:4]
curtimestr = curtime.split("T")[0].split(curyear + "-")[1]

# Logistic function definition
def logistic_func(x, L, k, x0):
    """Logistic function: L / (1 + exp(-k*(x - x0)))"""
    return L / (1 + np.exp(-k * (x - x0)))

fn = CONFIG["delegate-registration-spreadsheet-path"]

df = pd.read_excel(Path(fn))

dates: list[str] = list(df["Start time"])
nums: list[int] = [-1] * 1000
cur = 0

for i in range(len(dates)):
    dates[i] = str(dates[i]).split(" ")[0].split(curyear + "-")[-1]
    date = int(dates[i].split("-")[0]) * 31 + int(dates[i].split("-")[1]) - 53
    cur += 1
    nums[date] = cur

curdate = curtimestr
print(curdate)
date = int(curdate.split("-")[0]) * 31 + int(curdate.split("-")[1]) - 53
nums[date] = cur

x, y = [], []
diff = [0]
firstidx = -1
for i, n in enumerate(nums):
    if n == -1:
        continue
    if firstidx == -1:
        firstidx = i
    x.append(i)
    y.append(n)
    if len(y) > 1:
        diff.append(y[-1] - y[-2])
        
print(f"Current Total: {cur}")
diff[0] = nums[firstidx]

# Fit logistic regression
x_data = np.array(x)
y_data = np.array(y)

# Initial parameter guesses:
# L: maximum value (carrying capacity), use 80 as initial guess to match ylim
# k: growth rate, start with a small positive value
# x0: midpoint, guess based on data range
initial_guess = [80, 0.1, np.median(x_data)]

# Try to fit logistic curve, if it fails, use sensible defaults
try:
    params, covariance = curve_fit(logistic_func, x_data, y_data, 
                                   p0=initial_guess, maxfev=5000)
    L_fit, k_fit, x0_fit = params
    
    # Generate smooth logistic curve for the full x-range (0 to 60)
    x_smooth = np.linspace(0, 60, 200)
    y_logistic = logistic_func(x_smooth, L_fit, k_fit, x0_fit)
    
    # Plot logistic regression
    # plt.plot(x_smooth, y_logistic, 'r-', linewidth=2, 
    #          label=f'Logistic Regression\nL={L_fit:.1f}, k={k_fit:.3f}, x0={x0_fit:.1f}')
    
    print(f"Logistic parameters: L={L_fit:.1f}, k={k_fit:.3f}, x0={x0_fit:.1f}")
    
except (RuntimeError, ValueError) as e:
    print(f"Could not fit logistic curve: {e}")
    print("Using fallback logistic curve based on data range")
    
    # Fallback: simple logistic that extends to x=60, y≈80
    L_fallback = 80
    k_fallback = 0.2
    x0_fallback = np.mean(x_data)
    
    # x_smooth = np.linspace(0, 60, 200)
    # y_logistic = logistic_func(x_smooth, L_fallback, k_fallback, x0_fallback)
    
    # plt.plot(x_smooth, y_logistic, 'r--', linewidth=2, 
    #          label=f'Logistic Estimate\nL={L_fallback:.1f}, k={k_fallback:.3f}')

plt.bar(x, diff, label="Number of registrations at that day", alpha=0.7, width=0.5, color="orange")
plt.step(x, y, label="Cumulative Registration by Day", where="post")
plt.ylim((0, cur + 1))
# plt.xlim((0, 60))
plt.legend()
plt.xlabel("Day")
plt.ylabel("Number")
plt.title("Cumulative Registration by Day")

plt.show()