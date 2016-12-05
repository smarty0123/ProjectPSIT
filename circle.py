import matplotlib.pyplot as plt

def pic():
    slices = [7, 2, 2, 13] #Data
    activities = ['sleeping', 'eating', 'working', 'playing']
    cols = ['red', 'yellow', 'green', 'blue', 'pink']

    plt.pie(slices, labels = activities, colors = cols, startangle = 90, autopct = '%1.1f%%')
    plt.title('GRAPH')
    plt.show()
