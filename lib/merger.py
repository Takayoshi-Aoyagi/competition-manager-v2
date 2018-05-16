# -*- coding: utf-8 -*-

def merge_if_needed(entrant, merge_configs):
    for conf in merge_configs:
        ev = conf['event']
        if entrant[ev] not in conf['names']:
            continue
        entrant[ev] = conf['new_name']
    
class Marger:

    @staticmethod
    def merge(entrants, merge_configs):
        for entrant in entrants:
            merge_if_needed(entrant, merge_configs)
        return entrants
            
            
