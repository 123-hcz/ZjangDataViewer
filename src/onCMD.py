import click #导入库click

@click.command()#设置命令
@click.option('--count', default=2, help='Number of greetings. 问候的次数。')
@click.option('--name', prompt='Your name',
              help='The person to greet. 要问候的人。')
def main(count, name):#主函数
    #"""你可以输入以下命令."""
    """$一个简单的程序，它向这个名字打招呼，表示总的计数次数
        | Simple program that greets NAME for a total of COUNT times
    """
    #此处的"""  """用户查询命令时会输出，出来。
    for x in range(count):
        click.echo('Hello %s!' % name)

if __name__ == '__main__':
    main()