from pathlib import Path
import pypandoc
import pandas as pd
import regex as re
import time
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer
from datetime import datetime
from striprtf.striprtf import rtf_to_text


class FactivaParse:
    
    def __init__(self, *, path = "./projects"):
        """Initialises the project. 
- path --> a local directory that contains folders or RTF files.
           If this isn't present, it'll try again until 
           provided with an empty response."""
        temp_path = path
        while not Path(temp_path).is_dir():
            print(f"Directory '{temp_path}' not found. \n"
                  "Enter the relative path to the project library "
                  "or press return to close.")
            temp_path = input(":> ")
            if str(temp_path) == "":
                break
        if not temp_path:
            raise ValueError("Project did not complete initialisation")
        self._initialised = None
        self._file_dir = dict()
        self._data_store = Path(temp_path)

        print(f"Target folder set to: {self._data_store}")
        
        self._refactor_project_directories()
    
    
    ##
    ## Interfaces for project directories
    ##
    
    def _refactor_project_directories(self):
        self._file_dir = self._proc_candidate_folders_for_rtfs()
        
        # if the project hasn't been init'd then work silently
        # otherwise give a warning
        self._refactor_ids(warn=bool(self._initialised))
        
        # reset the initialisation time
        self._initialised = datetime.now()
        
    
    def _refactor_ids(self, *, force=False, warn=True):
        "refactor IDs for the project folders"
        id_ = 0
        for entry in self._file_dir:
            self._file_dir[entry]['id'] = id_
            id_+=1
        if warn:
            print(f"Project IDs refactored, {id_} IDs sorted.")
    
    
    def _proc_candidate_folders_for_rtfs(self):
        file_dir = dict()
        id_ = 0
        for _dir in self._data_store.glob("*"):
            if _dir.is_dir():
                _dir_name = _dir.parts[-1]
                file_dir[_dir_name] = {"id": id_, "proj": _dir_name, 
                                       "active": False, "files": list()}
                for file in _dir.glob("*.rtf"):
                    file_dir[_dir.parts[-1]]['files'].append(file)
                id_+=1
        return file_dir
    
    
    ##
    ## UX scripts for controlling projects
    ##
    
    def print_info(self):
        print(self._info_panel(head=True))

        
    def _run(self):
        "Run script from commandline"
        files = list(Path(".").rglob("*.rtf"))
        if len(files) ==0:
            raise ValueError("No valid target files. Run self.select_project")
        procced_files = self._proc_files(files, warn=True)
        _archive = procced_files['archive']
        DF = self._out_DF(archive)
        self._export_dataframe(DF=DF)

        
    def _gen_archive(self):
        files = self._get_active_project_files()
        if len(files) ==0:
            raise ValueError("No valid target files. Run self.select_project")
        procced_files = self._proc_files(files, warn=True)
        self._archive = procced_files['archive']
        self._search_archive = procced_files['search']
        
        
    def run(self):
        "Run script with human interface"
        self._select_project()
        self._gen_archive()
        self.DF = self._out_DF(archive)
        print("Dataframe created - self.DF")
        

    def _export_dataframe(self, *, DF=None):
        if type(DF) is type(None):
            DF = self.DF
        now = (datetime.now()).strftime("%Y-%m-%d_%I.%M%p")
        file_name = self._get_file_names()
        if len(file_name)>=24:
            file_name = file_name[0:24]
        else:
            file_name = f"{file_name}_{now}"
        print(f"Exporting spreadsheet '{file_name}.xlsx'")
        DF.to_excel(f"{file_name}.xlsx", engine='xlsxwriter')

        
    def _select_project(self):
        "interface for toggling projects on and off"
        self._refactor_ids(warn=False)
        finished = False
        print(self._info_panel())
        print("\n---------------------\n\n")
        print("Commands:\n\t- provide a project ID;"
              "\n\t- Type '[a]ll' to set all projects on;"
              "\n\t- Type '[n]one' to set all projects off;"
              "\n\t- Type '[i]nfo' to show project status;"
              "\n\t- Hit enter to conclude.")
        time.sleep(0.2)
        commands = ["all", "a", "n", "none", "i", "info"]
        
        while not finished:
            _id = input("\n:> ")
            if _id:
                if any(_id.lower() == command for command in commands):
                    # if val is 'a', then we're doing all to True
                    # otherwise sets everything to false
                    if _id.lower()[0] == "i":
                        print(self._info_panel())
                    else:
                        val = (_id.lower()[0] == "a")
                        for k, v in self._file_dir.items():
                            v['active'] = val
                        if val:
                            print("~~> All projects set 'ON'")
                        else:
                            print("~~> All projects set 'OFF'")
                else:
                    try:
                        _id = int(_id.strip())
                        try:
                            self._toggle_project(_id)
                        except ValueError as e:
                            print(e)
                    except ValueError:
                        print("Please provide a valid number.")
            else: # i.e. if '_id' is empty
                finished = True
        print("---------------------\n"
              "------CONCLUDING-----\n"
              "---------------------\n")
        print("The following projects are active:\n")
        print(self._info_panel(head=True, _all=False))
    
        
    def _get_proj_by_id(self, _id, *, warn=True):
        if not isinstance(_id, int):
            try:
                _id = int(_id.strip())
            except ValueError:
                raise ValueError(f"Provide an integer value. "
                                 f"Type {type(_id)} provided")
        for entry in self._file_dir:
            if self._file_dir[entry]['id'] == _id:
                return self._file_dir[entry]
        if warn:
            print(f"No project with id '{_id}' found.")
        return None
        
    def _get_proj_by_dir(self, _dir, *, warn=True):
        val = self._file_dir.get(_dir)
        if not val and warn:
            print(f"No project with dir '{_dir}' found.")
        return val
    
    def _info_panel(self, *, target=None, head=True, _all=True):
        out_s = ""
        if target and target not in self._file_dir:
            raise FileNotFoundError(f"Project '{target}' not found")
        if head:
            out_s = (f"{out_s}   ID | ON | # Files | Directory \n")
            out_s = (f"{out_s}------+----+---------+------------ - - - - - -\n")
        if target in self._file_dir:
            out_s = f"{out_s}{self._directory_status(target)}\n"
        else:
            for dir_ in self._file_dir:
                if _all or self._file_dir[dir_]['active']:
                    out_s = f"{out_s}{self._directory_status(dir_)}\n"
        return out_s
    
    def _get_keys_for_active_projects(self):
        for dir_ in self._file_dir:
            if _all or self._file_dir[dir_]['active']:
                out_s = f"{out_s}{self._directory_status(dir_)}\n"
        
                
    def _directory_status(self, target_):
        dir_ = self._file_dir[target_]
        row = f"  {' ' * (3 - len(str(dir_['id'])))}{dir_['id']} |"
        if dir_['active']:
            row=f"{row} ✓✓ |"
        else:
            row=f"{row}    |"
        num_files = str(len(dir_['files']))
        num_files = " " * (7 - len(num_files)) + num_files
        row = (f"{row} {num_files} | {self._data_store}/{target_}")
        return row
            
    def _toggle_project(self, dir_):
        """Given value 'dir_', toggles the entry True to False or vice versa.
If an integer is provided, then it will attempt to identify by ID number."""
        if isinstance(dir_, int):
            for entry in self._file_dir:
                if self._file_dir[entry]['id'] == dir_:
                    target=entry
                    break
            else:
                raise ValueError(f"Numeric identifier '{dir_}' "
                                 "not found in the file archive.")
        else:
            if dir_ in self._file_dir:
                entry = dir_
            else:
                raise ValueError(f"Entry '{dir_}' not found in the file archive.")
        self.__toggle_project_by_key(entry)
        
    
    def __toggle_project_by_key(self, entry, *, alert=True):
        temp_dir = self._file_dir[entry]
        temp_dir['active'] = not temp_dir['active']
        if alert:
            print(f"~~> Project {temp_dir['id']} -- {temp_dir['proj']} "
                  f"-- set to {temp_dir['active']}")

    def _string_right_align(i_str, indent=4, ideal_line_length = 60, max_line_length=70, lead_indent=4):
        tracker = lead_indent
        words = i_str.split()
        o_str = ""
        first = True
        for word in words:
            tracker+=len(word)
            if first:
                first = False
                o_str = o_str + word
            else:
                if tracker>ideal_line_length:
                    o_str = o_str + "\n"
                    tracker= 0
                if o_str[-1] == "\n":
                    o_str = o_str + " " * indent + word
                else:
                    o_str = o_str + " " + word
        return o_str
            
            
    
    ##
    ## Various properties
    ##
    
    @property
    def initialised(self):
        return self._initalised
    
    def __repr__(self):
        out_s = self._info_panel(_all=False)
        return out_s
    
    def __str__(self):
        out_s = self._info_panel(_all=False)
        return out_s
    
    def _get_active_project_files(self):
        o_files = list()
        for k, v in self._file_dir.items():
            if v['active']:
                o_files.extend(v['files'])
        return o_files

    def _get_file_names(self):
        f_names = list()
        for k, v in self._file_dir.items():
            if v['active']:
                return(v['proj'])
        return "DEFAULT"
    
    
    def _gen_archive(self):
        files = self._get_active_project_files()
        if len(files) ==0:
            raise ValueError("No valid target files. Run self.select_project")
        procced_files = self._proc_files(files, warn=True)
        self._archive = procced_files['archive']
        self._search_archive = procced_files['search']
    
    
    def _proc_files(self, files, *, warn=True):
        hl_re = re.compile("HYPERLINK(.*?)\\\\f28 \\\\par")
        count = 1
        archive = list()
        search_archive = list()
        if warn:
            print("Generating archive")
        for hl_rtf in files:
            if warn:
                print(f" - iterating file {count} {hl_rtf}")
            with open(hl_rtf) as file:
                corpus = file.read()
                search = Search(rtf_corpus=corpus)
                # will skip if something that doesn't resemble a Factiva search
                # is provided

                if search:
                    articles = hl_re.findall(corpus)
                    search_archive.append(search)
                    for article in articles:
                        archive.append(Article(article, search_hash=search.hash))
                elif warn:
                    print(f" - skipping file '{hl_rtf}' - not a Factiva file.")
            if warn:
                print(f" - archive now {len(archive)} articles")
            count+=1
        if warn:
            print("\n---------------------\n"
                  "------CONCLUDING-----\n"
                  "---------------------\n")
            print(f"Archive created with {len(archive)} articles")
        return {"archive": archive, "search": search_archive}
        
    def _out_DF(self, archive):
        o_list = list()
        for datapoint in archive:
            try:
                o_list.append(datapoint.data)
            except AttributeError:
                pass
        o_dataframe = pd.DataFrame(o_list)
        return o_dataframe
    
    
class Search:
    
    def __init__(self, rtf_corpus, strict=True):
        try:
            self._search = self._search_summary(rtf_corpus, strict=True)
        except:
            self._search = {}
        
    
    @property
    def search(self):
        return self._search
    
    
    def __hash__(self):
        """While the dict is hashable, the hash value would misidentify identical searches 
        as different due to the way that Factiva provides a different datetime instance 
        as a user page through search results. The fix is to directly target the calculated hash."""
        return self.search['hash']

    @property
    def hash(self):
        return self._search['hash']
    
    
    def __eq__(self, other):
        if type(self) == type(other) and self.hash == other.hash:
            return True
        else:
            return False
    
    
    def __str__(self):
        o_str = "Search("
        for k, v in self.search.items():
            if k[0:2] != "c_":
                o_str = f"{o_str}\n'{k}': {v}"
        return o_str+")"
    
    
    def __repr__(self):
        o_str = f"Search({self.search})"
        return o_str
    
    def __getitem__(self, key):
        return self._search[key]
    
    def _search_summary(self, rtf_corpus, strict=True):
        search_dict = dict()
        for param in rtf_to_text(rtf_corpus).split('Search Summary')[-1].split('\n'):
            if "|" in param:
                quasi_dict = param.strip().split('|')
                search_dict[quasi_dict[0]] = "".join(quasi_dict[1:]).replace("\u200b", "")
        hash_keys = ['Text', 'Date', 'Source', 'Author',
                     'Company', 'Subject', 'Industry', 'Region',
                     'Language', 'Results Found']
        # Timestamp is excluded from hash_keys, because it varies per page.
        hash_str = ""
        for key in hash_keys:
            hash_str = f"{hash_str}{search_dict[key]}"
        search_dict['hash'] = hash(hash_str)

        # patch dates
        date = search_dict['Date']
        start, end = None, None
        try:
            start, end = date.split(" to ", 1)
            try:
                start = datetime.strptime(start, "%d/%m/%Y")
            except ValueError:
                pass
            try:
                end = datetime.strptime(end, "%d/%m/%Y")
            except ValueError:
                pass
        except ValueError:
            pass
        search_dict['c_date'] = {'start': start, 'end':end, 'text':date}
        search_dict['c_results'] = "".join(re.findall('\d', search_dict['Results Found']))
        search_dict['c_timestamp'] = datetime.strptime(search_dict['Timestamp'], "%d %B %Y %H:%M")
        #patch results
        return search_dict

class Article:
    
    def __init__(self, txt_str, *, search_hash=None):
        read_rtf = rtf_to_text("{{{" + txt_str)
        
        date_re = re.compile("\d\d? \w* \d\d\d\d")
        author_re = re.compile("words?,(.*?)\(")
        words_re = re.compile("(\d*) \\\\uc2 word")
        time_re = re.compile("\D(\d\d?\:\d\d)\D")
        
        self.data = dict()
        self.raw = txt_str
        self.read_rtf = read_rtf
        self.search_hash = search_hash
        
        sample_txt = txt_str.split("{")
        url = sample_txt[0].split('\"')[1]
        title = rtf_to_text(txt_str.split("{")[-1].split("\\uc2")[1]).strip()
        source_and_time = sample_txt[-1].split("\\uc2")[2].strip()
        if (time:=time_re.search(source_and_time)):
            time = time[1]
        source = "".join(time_re.split(source_and_time))
        while source[-1] == " " or source[-1] == ",":
            source = source[0:-1]

        date = date_re.search(txt_str)[0]
        wordcount = words_re.search(txt_str)[1]
        if wordcount:
            try:
                wordcount = int(wordcount)
            except ValueError:
                pass
        
        authors = list()
        for can_auth in author_re.findall(read_rtf):
            
            for auth in re.split(" AND | AND,| BY |\,|\&", can_auth.upper(), flags=re.IGNORECASE):
                if auth:
                    authors.append(auth)
        authors = ", ".join(authors)
        
        if time:
            d_t = f"{time} {date}"
            date_time = datetime.strptime(d_t, "%H:%M %d %B %Y")
        else:
            date_time = None
        date = datetime.strptime(date, "%d %B %Y").date()
        
        self.data['url'] = url
        self.data['title'] = title
        self.data['source'] = source
        self.data['date'] = date
        self.data['datetime'] = date_time
        self.data['wordcount'] = wordcount
        self.data['authors'] = authors
        self.data['proxy_url'] = "redir".join(url.split("redir", 1)[1:])
        self.data.update(self.sentiment(title))
        self.data['lede'] = "\n".join(read_rtf.split('\n', 2)[2:])
        
    @staticmethod
    def sentiment(vals):
        analyser = SentimentIntensityAnalyzer()
        score = analyser.polarity_scores(vals)
        return score
        
    def __repr__(self):
        out_s = self.data['title']
        for k,v in self.data.items():
            if k == 'title':
                pass
            else:
                out_s = (f"{out_s}\n  {k}\t{v}")
        return out_s

if __name__=="__main__":
    FactivaParse()._run()
