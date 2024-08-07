## code to test matplotlib

import matplotlib.pyplot as plt

# Sample data
x = [1, 2, 3, 4, 5]
y = [2, 3, 5, 7, 11]

# Create a plot
plt.plot(x, y, marker='o')
plt.title('Sample Plot')
plt.xlabel('X Axis')
plt.ylabel('Y Axis')
plt.grid(True)

# Show the plot
plt.show()