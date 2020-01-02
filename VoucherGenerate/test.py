from progressbar import Percentage, Bar, Timer, ETA, progressbar


progress = progressbar.ProgressBar(widgets=['Progress: ', Percentage(), ' ', Bar('#'), ' ', Timer(),
                                            ' ', ETA(), ' ', ' '])

l = []
for i in progress(range(10000000)):
    l.append(1)
