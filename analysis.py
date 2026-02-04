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

L_fit: float | int =0
L_fallback: float | int =0

# Try to fit logistic curve, if it fails, use sensible defaults
try:
    params, covariance = curve_fit(logistic_func, x_data, y_data, 
                                   p0=initial_guess, maxfev=5000)
    L_fit, k_fit, x0_fit = params
    # Generate smooth logistic curve for the full x-range (0 to 60)
    x_smooth = np.linspace(0, 60, 200)
    y_logistic = logistic_func(x_smooth, L_fit, k_fit, x0_fit)
    
    # Compute R^2 on the observed x-data using the fitted parameters
    y_pred_fit = logistic_func(x_data, L_fit, k_fit, x0_fit)
    ss_res = np.sum((y_data - y_pred_fit) ** 2)
    ss_tot = np.sum((y_data - np.mean(y_data)) ** 2)
    if ss_tot == 0:
        r2 = 1.0 if ss_res == 0 else 0.0
    else:
        r2 = 1.0 - (ss_res / ss_tot)

    # Plot logistic regression with R^2 in the legend
    plt.plot(x_smooth, y_logistic, 'r-', linewidth=2, 
             label=f'Logistic Regression\nL={L_fit:.1f}, k={k_fit:.3f}, x0={x0_fit:.1f}, R2={r2:.3f}')
    
    print(f"Logistic parameters: L={L_fit:.1f}, k={k_fit:.3f}, x0={x0_fit:.1f}, R^2={r2:.3f}")
    
except (RuntimeError, ValueError) as e:
    print(f"Could not fit logistic curve: {e}")
    print("Using fallback logistic curve based on data range")
    
    # Fallback: simple logistic that extends to x=60, y≈80
    L_fallback = 80
    k_fallback = 0.2
    x0_fallback = np.mean(x_data)
    
    x_smooth = np.linspace(0, 60, 200)
    y_logistic = logistic_func(x_smooth, L_fallback, k_fallback, x0_fallback)
    
    # Compute R^2 for the fallback curve as an informative estimate
    y_pred_fallback = logistic_func(x_data, L_fallback, k_fallback, x0_fallback)
    ss_res_fb = np.sum((y_data - y_pred_fallback) ** 2)
    ss_tot_fb = np.sum((y_data - np.mean(y_data)) ** 2)
    if ss_tot_fb == 0:
        r2_fb = 1.0 if ss_res_fb == 0 else 0.0
    else:
        r2_fb = 1.0 - (ss_res_fb / ss_tot_fb)

    plt.plot(x_smooth, y_logistic, 'r--', linewidth=2, 
             label=f'Logistic Estimate\nL={L_fallback:.1f}, k={k_fallback:.3f}, R2={r2_fb:.3f}')
    print(f"Using fallback logistic curve. R^2={r2_fb:.3f}")

upper_lim = max(L_fit or L_fallback or 0, cur) + 1 # type: ignore

plt.bar(x, diff, label="Number of registrations at that day", alpha=0.7, width=0.5, color="orange")
plt.step(x, y, label="Cumulative Registration by Day", where="post")
plt.ylim((0, upper_lim))
# plt.xlim((0, 60))
plt.legend()
plt.xlabel("Day")
plt.ylabel("Number")
plt.title("Cumulative Registration by Day")

plt.show()