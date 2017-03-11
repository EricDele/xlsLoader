#! /usr/bin/env python
# -*- coding: utf-8 -*-
import loader as xls

if __name__ == '__main__':
    # Instanciate the object
    my_data = xls.Loader('conf/config.json')
    # Load the xls file in the object
    my_data.load_file('../YarnQueueManager/xls/Implementation_Queues_PROD-YARN_1-2.xlsm')
    # Print the datas
    print(my_data)
    # Get the line numbers
    print(my_data.get_line_numbers())
    # print a sepcific line
    print(my_data.get_a_line(15))
    print("")
    # Iterate the datas
    for i in my_data:
        print(i)
    # write a csv file
    my_data.write_csv_file('test.csv')