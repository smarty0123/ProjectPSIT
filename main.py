import circle as c
import bar as b
def main():
    print('-Graph1    Show .....')
    print('-Graph2    Show .....')
    print('-Graph3    Show .....')
    print('-Graph4    Show .....')
    print('-Exit      Close Program')
    while True:
        print('PLEASE ENTER INPUT : ', end='')
        text = input()
        if text == 'Graph1':
            b.pic()
        elif text == 'Graph2':
            c.pic()
        elif text == 'Exit':
            break
main()
