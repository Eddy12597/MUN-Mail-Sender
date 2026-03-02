import pandas as pd
from pathlib import Path
import matplotlib.pyplot as plt
import numpy as np
from scipy.optimize import curve_fit
from config import *
import datetime

# 1. Setup Dates
# Start date: January 23, 2026
START_DATE = datetime.date(2026, 1, 23)
curtime = datetime.datetime.now(datetime.UTC)
curdate_obj = curtime.date()

# Define the function for fitting
def logistic_func(x, L, k, x0):
    """Logistic function: L / (1 + exp(-k*(x - x0)))"""
    return L / (1 + np.exp(-k * (x - x0)))

# 2. Data Loading
fn = CONFIG["delegate-registration-spreadsheet-path"]
df = pd.read_excel(Path(fn))

# 3. Proper Day Indexing
# Convert 'Start time' to datetime objects, handling potential string/timestamp mix
df["Start time"] = pd.to_datetime(df["Start time"])

# Calculate day index: (Registration Date - Jan 23)
# result.days gives us an integer 0, 1, 2...
day_indices = [(d.date() - START_DATE).days for d in df["Start time"]]

# We use a dictionary to count registrations per day to avoid fixed-size list issues
daily_counts = {}
for day in day_indices:
    daily_counts[day] = daily_counts.get(day, 0) + 1

# Ensure current date is represented in our range
max_day = max(max(daily_counts.keys()), (curdate_obj - START_DATE).days)
x_data = []
y_data = []
diff = []
cumulative = 0

# Build continuous timeline from Day 0 to max_day
for d in range(max_day + 1):
    count = daily_counts.get(d, 0)
    cumulative += count
    
    # We only plot points where we actually have data or at the very end
    # to keep the regression clean, but you can also include all days.
    if d in daily_counts or d == max_day:
        x_data.append(d)
        y_data.append(cumulative)
        diff.append(count)

x_data = np.array(x_data)
y_data = np.array(y_data)

print(f"Current Total: {cumulative}")
print(f"Days since start: {max_day}")

# 4. Fit Logistic Regression
initial_guess = [max(y_data) * 1.5, 0.1, np.median(x_data)]
L_plot, k_plot, x0_plot = 0, 0, 0
is_fit = False

try:
    params, _ = curve_fit(logistic_func, x_data, y_data, p0=initial_guess, maxfev=5000)
    L_plot, k_plot, x0_plot = params
    
    y_pred = logistic_func(x_data, L_plot, k_plot, x0_plot)
    ss_res = np.sum((y_data - y_pred) ** 2)
    ss_tot = np.sum((y_data - np.mean(y_data)) ** 2)
    r2 = 1.0 - (ss_res / ss_tot) if ss_tot != 0 else 0
    
    label_str = f'Logistic Fit: L={L_plot:.1f}, R2={r2:.3f}'
    is_fit = True
    
except Exception as e:
    print(f"Fit failed: {e}")
    # Fallback params
    L_plot, k_plot, x0_plot = max(y_data), 0.1, np.median(x_data)
    label_str = 'Logistic Estimate (Fallback)'

# 5. Plotting
plt.figure(figsize=(10, 6))

# Smooth curve for prediction
x_smooth = np.linspace(0, max(max_day, 60), 200)
y_smooth = logistic_func(x_smooth, L_plot, k_plot, x0_plot)

plt.bar(x_data, diff, label="Daily Registrations", alpha=0.4, color="orange")
plt.step(x_data, y_data, label="Cumulative Count", where="post", linewidth=2)
plt.plot(x_smooth, y_smooth, 'r--', label=label_str)

plt.ylim(0, max(L_plot * 1.1, cumulative + 10))
plt.xlabel(f"Days since Jan 23")
plt.ylabel("Number of Delegates")
plt.title("Delegate Registration Growth")
plt.legend()
plt.grid(axis='y', linestyle='--', alpha=0.7)

plt.show()