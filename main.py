from mailmerge import MailMerge
import random

numbers = [72,102,129,121,136,150,158,174,186,197,200,202,214,216,227,237,238,243,270,273,247,283,286,300,301,303,316,327,328,334,336,342,356,360,376,382,380,398,428,535,536]
n = 120


with MailMerge('template.docx') as document:
        # search google for merge field insertion in word
        dicts_random = []
        for _ in range(1,n+1):
            random_items = random.sample(numbers, 25)
            dicts_random.append({ str(i+1):str(random_items[i]) for i in range(25) })
        document.merge_templates(dicts_random, separator='page_break')
        document.write(f"bingo.docx")
        
