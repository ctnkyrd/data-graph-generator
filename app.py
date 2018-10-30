# -*- coding: utf-8 -*-
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import style
import pandas as pd

style.use('ggplot')

# create plot
fig, ax = plt.subplots()
bar_width = 0.35
opacity = 0.8
 
city=[u'Basılı Veri',u'Dijital Veri']
pos = np.arange(len(city))
y=[1, 23]
 
plt.bar(pos,y, bar_width,label=u"Veri Türü", alpha=opacity, color='blue',edgecolor='black')
plt.xticks(pos, city)
# plt.xlabel(u'Veri Türü', fontsize=16)
# plt.ylabel('Happiness_Index', fontsize=16)
plt.title(u'Veri Türü',fontsize=16)

plt.legend()
plt.tight_layout()
plt.savefig('testplot.png')



def autolabel(rects, ax):
    # Get y-axis height to calculate label position from.
    (y_bottom, y_top) = ax.get_ylim()
    y_height = y_top - y_bottom

    for rect in rects:
        height = rect.get_height()

        # Fraction of axis height taken up by this rectangle
        p_height = (height / y_height)

        # If we can fit the label above the column, do that;
        # otherwise, put it inside the column.
        if p_height > 0.95: # arbitrary; 95% looked good to me.
            label_position = height - (y_height * 0.05)
        else:
            label_position = height + (y_height * 0.01)

        ax.text(rect.get_x() + rect.get_width()/2., label_position,
                '%d' % int(height),
                ha='center', va='bottom')