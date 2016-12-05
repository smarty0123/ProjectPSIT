import matplotlib.pyplot as plt

def pic():

    x = [1, 2, 3, 4, 5]
    y = [4, 8, 6, 2, 7]

    name = ['SMART', 'JJ', 'MAX' , 'KHAN', 'EARTH']
    plt.bar(x, y, label = 'Bars1', align = 'center')

    plt.xticks(x, name, rotation='vertical')

    plt.xlabel('x')
    plt.ylabel('y')
    plt.legend()
    plt.show()
