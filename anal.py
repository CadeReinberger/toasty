import pandas as pd

SOURCE_FILE = 'SCORESHEET.xlsx'
TARGET_FILE = 'STATS.xlsx'

class writeline:
    def __init__(self, written):
        self.written = written
        
    def to_str(self):
        return 'W:' + str(self.written)
    
    def to_dict(self):
        return {'questions' : self.written,
                'scoreline' : self.to_str(),
                'ppg' : 200.00}
    
DIDNT_WRITE = writeline(0)

class statline:
    def __init__(self, powers, tossups, negs, questions):
        self.questions = questions
        self.powers = powers
        self.tossups = tossups
        self.negs = negs        
        
    def to_str(self):
        return str(self.powers) + r'/' + str(self.tossups) + '/' + str(self.negs)
    
    def ppg(self):
        return 20*(15*self.powers + 10*self.tossups - 5*self.negs)/self.questions
    
    def to_dict(self):
        return {'questions' : self.questions,
                'scoreline' : self.to_str(),
                'ppg' : self.ppg()}
    
DIDNT_PLAY = statline(0, 0, 0, 0)
    
class player:
    def __init__(self):
        self.combined_stats = statline(0, 0, 0, 0)
        self.writing_stats = writeline(0)
        
    def update(self, line):
        if isinstance(line, writeline):
            self.writing_stats = line
        else:
            self.combined_stats.questions += line.questions
            self.combined_stats.powers += line.powers
            self.combined_stats.tossups += line.tossups
            self.combined_stats.negs += line.negs

    def to_str(self):
        return self.combined_stats.to_str() + '/' + self.writing_stats.to_str()
    
    def tot_points(self):
        writ = self.writing_stats.written
        pwr = self.combined_stats.powers
        tus = self.combined_stats.tossups
        neg = self.combined_stats.negs
        return 10*writ + 15*pwr + 10*tus - 5*neg
    
    def to_dict(self):
        return {"questions" : self.combined_stats.questions,
                "statline" : self.to_str(),
                "ppg" : self.combined_stats.ppg(),
                "points" : self.tot_points()}
    
#get the relevant stuff
writers = pd.ExcelFile(SOURCE_FILE).sheet_names 

def read_writer(writer):
    df = pd.read_excel(SOURCE_FILE, writer)
    res = {}

    #strip any spaces so that an extraneous space on an entry doesn't make a new player in the stats
    df['Answerer'] = df['Answerer'].str.strip()
    #make all answerer names title case so typos in capitalization don't make a new player in the stats
    df['Answerer'] = df['Answerer'].str.title()
    
    #get number of questions
    qs = []
    for q in df['Question']:
        try:
            qs.append(int(q))
        except:
            pass
    questions = max(qs)
    
    for ind in range(len(df)):
        player = df['Answerer'].iloc[ind]
        if player == 'NaN':
            continue
        if player not in res:
            res[player] = statline(0, 0, 0, questions)
            
        try:
            score = int(df['Points'].iloc[ind])
            if score == -5:
                res[player].negs = res[player].negs + 1
            elif score == 10:
                res[player].tossups = res[player].tossups + 1
            elif score == 15:
                res[player].powers = res[player].powers + 1
        except: 
            pass
    
    res[writer] = writeline(questions)
    
    return res
        
    
def make_writer_summary(round_data):
    resdicts = { r : c.to_dict() for (r, c) in round_data.items()}
    resdf = pd.DataFrame.from_dict(resdicts).T
    findf = resdf.sort_values(by = 'ppg', ascending = False)
    return findf
    
player_res = {}
round_res = {}

for writer in writers:
    rd_data = read_writer(writer)
    round_res[writer] = make_writer_summary(rd_data)
    
    for plr in rd_data:
        if plr not in player_res:
            player_res[plr] = player()
        player_res[plr].update(rd_data[plr])
        
plyrdicts = { r : c.to_dict() for (r, c) in player_res.items()}
player_res_dict = pd.DataFrame.from_dict(plyrdicts).T
res_data = player_res_dict.sort_values(by = 'points', ascending = False)    

#finally, we just output the data
with pd.ExcelWriter(TARGET_FILE) as writer:  
    res_data.to_excel(writer, sheet_name='FINAL')
    for (name, df) in round_res.items():
        df.to_excel(writer, sheet_name=name)
    
    

