#!/usr/bin/env python3
"""
AI Presentation Agent — Live Web Search Integration
===================================================
New: Real-time web search tool for local models. 
Uses duckduckgo-search (no API key required) to ground presentations in real facts.
"""

import sys, os, json, re, time, textwrap, platform, subprocess
from pathlib import Path
from dataclasses import dataclass, field, asdict
from typing import List, Optional, Dict, Any, Callable, Tuple
from functools import wraps

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QTextEdit, QComboBox, QSpinBox, QPushButton,
    QTabWidget, QScrollArea, QFrame, QFileDialog, QMessageBox,
    QCheckBox, QStackedWidget, QSplitter, QStatusBar, QGridLayout,
    QSlider, QSizePolicy, QListWidget
)
from PySide6.QtCore import Qt, Signal, QThread, QTimer, QSize
from PySide6.QtGui import QFont, QColor, QPalette, QCursor, QPainter

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

try: import openai; HAS_OPENAI = True
except ImportError: HAS_OPENAI = False
try:
    from reportlab.lib.pagesizes import LETTER, landscape
    from reportlab.lib.colors import HexColor, white, black
    from reportlab.pdfgen import canvas
    HAS_REPORTLAB = True
except ImportError: HAS_REPORTLAB = False
try: import requests; HAS_REQUESTS = True
except ImportError: HAS_REQUESTS = False

# IMPROVEMENT: Web Search
try:
    from duckduckgo_search import DDGS; HAS_SEARCH = True
except ImportError: HAS_SEARCH = False

C = {'base':'#1e1e2e','mantle':'#181825','crust':'#11111b','surface0':'#313244','surface1':'#45475a','surface2':'#585b70','overlay0':'#6c7086','text':'#cdd6f4','subtext':'#a6adc8','blue':'#89b4fa','green':'#a6e3a1','red':'#f38ba8','yellow':'#f9e2af','peach':'#fab387','mauve':'#cba6f7','teal':'#94e2d5','lavender':'#b4befe'}
TEMPLATES = {
    'executive_blue': {'name':'Executive Blue','icon':'🔷','bg':(0x0A,0x1F,0x44),'title_c':(0xFF,0xFF,0xFF),'body_c':(0xD0,0xD8,0xE8),'accent':(0x3B,0x82,0xF6),'title_font':'Calibri','body_font':'Calibri','title_sz':32,'body_sz':16,'bullet_sz':15},
    'dark_modern': {'name':'Dark Modern','icon':'◼️','bg':(0x1E,0x1E,0x2E),'title_c':(0xC9,0xD1,0xD9),'body_c':(0xA0,0xA8,0xB8),'accent':(0x89,0xB4,0xFA),'title_font':'Segoe UI','body_font':'Segoe UI','title_sz':30,'body_sz':15,'bullet_sz':14},
}

def robust_json_extract(text: str) -> Any:
    if not text: return None
    text = re.sub(r'^```(?:json)?\s*', '', text.strip()); text = re.sub(r'\s*```$', '', text).strip()
    target_char = '[' if '[' in text else '{'
    start = text.find(target_char)
    if start == -1: return None
    depth, end = 0, -1
    for i in range(start, len(text)):
        if text[i] == target_char: depth += 1
        elif (target_char == '[' and text[i] == ']') or (target_char == '{' and text[i] == '}'):
            depth -= 1
            if depth == 0: end = i; break
    if end == -1: return None
    try: return json.loads(re.sub(r',\s*([}\]])', r'\1', text[start:end+1]))
    except: return None

@dataclass
class SlideData:
    number:int=0; title:str=""; subtitle:str=""; layout:str="content"
    body:str=""; bullets:List[str]=field(default_factory=list)
    highlight:str=""; speaker_notes:List[str]=field(default_factory=list)
    full_script:str=""; timing:int=30; tips:str=""; visual_suggestion:str=""
    metadata:Dict[str,Any]=field(default_factory=dict)

@dataclass
class PresConfig:
    topic:str=""; audience:str="General professionals"; num_slides:int=10
    tone:str="Professional"; style:str="Informative"; template:str="executive_blue"
    include_notes:bool=True; include_script:bool=True; extra:str=""; system_override:str=""
    enable_search:bool=True # NEW: Web search toggle

@dataclass
class GPUConfig:
    n_gpu_layers:int=0; n_ctx:int=8192; n_threads:int=8; batch_size:int=512
    n_parallel:int=4; temperature:float=0.7; top_p:float=0.9; top_k:int=40
    repeat_penalty:float=1.1; seed:int=-1; memory_limit_gb:float=0.0
    model_path:str=""; server_host:str="127.0.0.1"; server_port:int=8080; server_cmd:str="llama-server"

# ═════════════════════════════════════════════════════════════
# TOOLING & AI BACKEND
# ═════════════════════════════════════════════════════════════
class ToolRegistry:
    def __init__(self): self._tools={}; self._log=[]
    def register(self, t): self._tools[t.name]=t
    def schemas(self): return [t['schema'] for t in self._tools.values()]
    def execute(self, name, args):
        t=self._tools.get(name)
        if t and t['handler']: 
            try: r=t['handler'](**args); self._log.append({'tool':name,'res':str(r)[:150]}); return r
            except Exception as e: return {"error":str(e)}
        return {"error":"Not found"}

class PresentationTools:
    def __init__(self, cfg, gpu=None):
        self.registry=ToolRegistry()
        self.registry.register({'name':'suggest_layout','schema':{"type":"function","function":{"name":"suggest_layout","description":"Suggest layout based on content goal","parameters":{"type":"object","properties":{"content_goal":{"type":"string","enum":["hook","narrative","data_dense","comparison","process","conclusion"]}}}},'handler':lambda content_goal: {"layout":"content" if content_goal in ["narrative","hook"] else "bullets" if content_goal=="data_dense" else "two_column" if content_goal=="comparison" else "bullets"}})
        self.registry.register({'name':'validate_fit','schema':{"type":"function","function":{"name":"validate_fit","description":"Validate slide content limits","parameters":{"type":"object","properties":{"title_len":{"type":"integer"},"body_len":{"type":"integer"},"bullet_count":{"type":"integer"},"longest_bullet":{"type":"integer"}}}},'handler':lambda title_len,body_len=0,bullet_count=0,longest_bullet=0: {"fits": title_len<=60 and body_len<=500 and bullet_count<=6 and longest_bullet<=90, "violations": [x for x in [("Title too long", title_len>60), ("Body too long", body_len>500), ("Too many bullets", bullet_count>6), ("Bullet too long", longest_bullet>90)] if x[1]]}})
        self.registry.register({'name':'calc_timing','schema':{"type":"function","function":{"name":"calc_timing","description":"Calculate pacing","parameters":{"type":"object","properties":{"layout":{"type":"string"},"complexity":{"type":"string","enum":["low","medium","high"]}}}},'handler':lambda layout, complexity="medium": {"seconds": 20 if layout=="title" else 30 if complexity=="low" else 45}}})
        
        # IMPROVEMENT: WEB SEARCH TOOL
        if HAS_SEARCH and cfg.enable_search:
            self.registry.register({
                'name': 'web_search',
                'schema': {
                    "type": "function",
                    "function": {
                        "name": "web_search",
                        "description": "Search the web for current, factual information. USE THIS when the topic involves recent events, specific statistics, real-time data, or post-2023 information. Do NOT use for general knowledge. Always cite sources.",
                        "parameters": {
                            "type": "object",
                            "properties": {
                                "query": {"type": "string", "description": "The exact search string to look up"}
                            },
                            "required": ["query"]
                        }
                    }
                },
                'handler': self._web_search_handler
            })

    def _web_search_handler(self, query: str, max_results: int = 5) -> Dict:
        """Executes a web search and returns formatted results."""
        try:
            results = []
            for r in DDGS(query, max_results=max_results):
                results.append({
                    "title": r.title,
                    "body": r.body if r.body else "",
                    "href": r.href,
                    "source": r.source if hasattr(r, 'source') else ""
                })
            if not results:
                return {"error": "No search results found. Proceed with internal knowledge."}
            return {"results": results, "count": len(results), "query": query}
        except Exception as e:
            return {"error": f"Search failed: {str(e)}. Proceed with internal knowledge."}

# ═════════════════════════════════════════════════════════════
# v2.4 EXECUTIVE ARCHITECT SKILLS
# ═════════════════════════════════════════════════════════════
class Skill:
    name="base"
    def sys(self,c): return ""
    def usr(self,c,sl=None): return ""
    def parse(self,r,ex=None): return ex or []

class OutlineSkill(Skill):
    name="architect_outline"
    def sys(self,c):
        override = c.system_override + "\n\n" if c.system_override else ""
        search_rule = ""
        if HAS_SEARCH and c.enable_search:
            search_rule = """
7. WEB SEARCH RULE (If enabled): If the topic involves current events, specific data, statistics, or anything that happened after your training cutoff, you MUST call 'web_search' BEFORE creating the outline. Inject the facts you find directly into the slide key_points. CITE YOUR SOURCES in the bullets (e.g., "Revenue grew 23% [Source Name]")."""
        
        return f"""{override}You are a McKinsey-level Presentation Architect.

RULES FOR EXCELLENCE:
1. MINTO PYRAMID: Start with the conclusion FIRST. Then provide supporting arguments.
2. NARRATIVE ARC: Hook -> Context -> Conflict/Tension -> Resolution -> Action.
3. LAYOUT VARIETY: Do NOT use the same layout more than twice in a row.
4. MINTO PYRAMID PRINCIPLE: Start with the conclusion first. Never build up to the point; state the point, then prove it.
5. MANDATORY TOOLS: Call 'suggest_layout' for every slide to ensure variety.
{search_rule}
OUTPUT: Strictly return a JSON array:
[{{"slide_number": 1, "title": "Max 6 words", "layout": "title|content|bullets|two_column", "key_points": ["Max 3 points per slide."], "sources": ["Source (if web search used)"]}}]"""

    def usr(self,c,sl=None):
        return f"""Design the architecture for a {c.num_slides}-slide presentation.
TARGET: {c.topic}
AUDIENCE: {c.audience} | STYLE: {c.style} | TONE: {c.tone}
{f"CONSTRAINTS: {c.extra}" if c.extra else ""}
Step 1: {'Call "web_search" if c.enable_search else 'Skip search.'} 
Step 2: Call 'suggest_layout' for {c.num_slides} slides.
Step 3: Output the JSON array."""

    def parse(self,r,ex=None):
        d=robust_json_extract(r)
        if not d: return [SlideData(number=1,title="Architecture Error",body=r[:300])]
        return [SlideData(number=x.get('slide_number',i+1),title=x.get('title',''),layout=x.get('layout','content'),bullets=x.get('key_points',[]),metadata={"sources":x.get('sources',[])}) for i,x in enumerate(d)]

class ContentSkill(Skill):
    name="visual_content_writer"
    def sys(self,c):
        override = c.system_override + "\n\n" if c.system_override else ""
        search_rule = ""
        if HAS_SEARCH and c.enable_search:
            search_rule = """
5. FACTUAL ACCURACY: If you found data via web_search during the outline phase, EMBED the exact statistics and facts into the bullets and body. Format as: "Revenue grew 23% in Q3 2024 [Source]" to add credibility."""
        return f"""{override}You are an Elite Visual Slide Designer.
ANTI-BORING RULES:
1. NEWSPAPER HEADLINE TEST: Titles must be 6 words or fewer.
2. ARMCHAIR TEST: A bullet point must be readable in 3 seconds.
3. VISUAL-FIRST: Every slide needs a 'visual_suggestion'.
4. SO WHAT? TEST: Every piece of data must pass the "So what?" test.
5. MANDATORY VALIDATION: Call 'validate_fit' for EVERY slide.
{search_rule}
OUTPUT FORMAT:
[{{"number":1,"title":"6 words max","subtitle":"Optional","body":"Max 50 words if layout is 'content'","bullets":["Verb-led, max 80 chars"],"highlight":"One killer stat","visual_suggestion":"Exact graphical layout","sources":["Source if applicable"]}}]"""

    def usr(self,c,sl=None):
        enriched = [{"n":s.number,"t":s.title,"l":s.layout,"kp":s.bullets,"src":s.metadata.get("sources",[]),"goal":s.metadata.get("goal","narrative")} for s in sl]
        return f"""Write visual content for "{c.topic}".
AUDIENCE: {c.audience} | TONE: {c.tONE if hasattr(c, 'tONE') else c.tone}
OUTLINE TO EXPAND:
{json.dumps(enriched_slides, indent=2)}
CRITICAL: For EVERY slide, call 'validate_fit'. Output JSON array."""

    def parse(self,r,ex=None):
        d=robust_json_extract(r); r_list=list(ex or [])
        if not d: return r_list
        for x in d:
            s=next((sl for sl in r_list if sl.number==x.get('number')),None)
            if s:
                s.title=x.get('title',s.title); s.subtitle=x.get('subtitle','')
                s.body=x.get('body',''); s.bullets=x.get('bullets',s.bullets)
                s.highlight=x.get('highlight',''); s.visual_suggestion=x.get('visual_suggestion','')
                if x.get('sources'): s.metadata['sources'] = x.get('sources',[])
        return r_list

class NotesSkill(Skill):
    name="stage_director_notes"
    def sys(self,c):
        override = c.system_override + "\n\n" if c.system_override else ""
        search_rule = ""
        if HAS_SEARCH and c.enable_search:
            search_rule = """
5. CONTEXTUAL ENRICHMENT: If slides contain statistics or facts from web_search, ensure the script mentions the source. Example: "According to Gartner's 2024 report..." """
        return f"""{override}You are a Premium Executive Speech Coach.
STRUCTURE OF GREAT NOTES: [Cue, Concept, Bridge]
STRUCTURE OF GREAT SCRIPTS: Conversational, short sentences, explicit transitions.
{search_rule}
OUTPUT FORMAT:
[{{"number":1,"speaker_notes":["[Cue]","Concept: ...","Bridge: ..."],"full_script":"Natural spoken words...","timing_secs":45,"delivery_tip":"...","sources":["..."]}}]"""

    def usr(self,c,sl=None):
        ctx = [{"n":s.number,"t":s.title,"prev": sl[i-1].title if i > 0 else "Start","next": sl[i+1].title if i < len(sl)-1 else "End","src": s.metadata.get("sources",[])} for i,s in enumerate(sl)]
        return f"""Write notes for "{c.topic}".
SLIDE SEQUENCE (for transitions):
{json.dumps(ctx, indent=2)}
INSTRUCTIONS:
1. Write verbatim script including transitions.
2. Cite sources in the script if web search was used.
3. Output JSON array."""

    def parse(self,r,ex=None):
        d=robust_json_extract(r); r_list=list(ex or [])
        if not d: return r_list
        for x in d:
            s=next((sl for sl in r_list if sl.number==x.get('number')),None)
            if s:
                s.speaker_notes=x.get('speaker_notes',[]); s.full_script=x.get('full_script','')
                s.timing=x.get('timing_secs',30); s.tips=x.get('delivery_tip','')
                if x.get('sources'): s.metadata['script_sources'] = x.get('sources',[])
        return r_list

class SkillHarness:
    def __init__(self,prov,tools,log_cb=None):
        self.prov=prov; self.tools=tools; self.log_cb=log_cb; self.log=[]
    def execute(self,skill,cfg,slides=None):
        try:
            msgs=[{"role":"system","content":skill.sys(cfg)},{"role":"user","content":skill.usr(cfg,slides)}]
            for _ in range(10):
                resp=self.prov.chat_with_tools(msgs,self.tools.schemas())
                if not resp: break
                msg=resp.get('message',{}); msgs.append(msg)
                if not msg.get('tool_calls'): return skill.parse(msg.get('content',''),slides)
                for tc in msg.get('tool_calls',[]):
                    fn=tc.get('function',{})
                    try: args=json.loads(fn.get('arguments','{}'))
                    except: args={}
                    if self.log_cb: self.log_cb(f"🔧 {fn.get('name')}({json.dumps(args)[:80]})")
                    res=self.tools.execute(fn.get('name',''),args)
                    msgs.append({"role":"tool","tool_call_id":tc.get('id',''),"content":json.dumps(res) if not isinstance(res,str) else res})
        except Exception as e:
            if self.log_cb: self.log_cb(f"⚠️ Tool loop failed, falling back...")
            try: return skill.parse(self.prov.chat_simple(skill.sys(cfg), skill.usr(cfg,slides)),slides)
            except: pass
        return slides or []

# ═════════════════════════════════════════════════════════════
# AI PROVIDERS
# ═════════════════════════════════════════════════════════════
def retry(max_r=2,delay=2.0):
    def dec(f):
        @wraps(f)
        def w(*a,**kw):
            for i in range(max_r):
                try: return f(*a,**kw)
                except: time.sleep(delay*(2**i))
        return w
    return dec

class AIProvider:
    def chat_with_tools(self,m,t): pass
    def chat_simple(self,s,u): pass

class LlamaCppProvider(AIProvider):
    def __init__(self,gpu): self.gpu=gpu; self.url=f"http://{gpu.server_host}:{gpu.server_port}"; self.proc=None
    @retry()
    def _post(self,ep,p,t=300): r=requests.post(f"{self.url}{ep}",json=p,timeout=t); r.raise_for_status(); return r.json()
    def health_check(self):
        try:
            if requests.get(f"{self.url}/health",timeout=2).status_code==200: return True,"OK"
        except: pass
        return False,"Offline"
    def start_server(self):
        if self.health_check()[0]: return True,"Running"
        if not os.path.exists(self.gpu.model_path): return False,"No model"
        cmd=[self.gpu.server_cmd,"-m",self.gpu.model_path,"--host",self.gpu.server_host,"--port",str(self.gpu.server_port),"-ngl",str(self.gpu.n_gpu_layers),"-c",str(self.gpu.n_ctx),"-t",str(self.gpu.n_threads)]
        try:
            self.proc=subprocess.Popen(cmd,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,creationflags=subprocess.CREATE_NO_WINDOW if platform.system()=='Windows' else 0)
            for _ in range(20):
                time.sleep(1)
                if self.health_check()[0]: return True,"Started"
        except Exception as e: return False,str(e)
        return False,"Timeout"
    def stop_server(self):
        if self.proc: self.proc.terminate(); self.proc.wait(timeout=5); self.proc=None
    @retry()
    def chat_with_tools(self,m,t): p={"messages":m,"max_tokens":4096,"tools":t,"tool_choice":"auto"}; return self._post("/v1/chat/completions",p)
    @retry()
    def chat_simple(self,s,u): return self._post("/v1/chat/completions",{"messages":[{"role":"system","content":s},{"role":"user","content":u}],"max_tokens":4096}).get('choices',[{}])[0].get('message',{}).get('content','')

class OpenAIProvider(AIProvider):
    def __init__(self,k,m="gpt-4o-mini",b=None): self.m=m; self.c=openai.OpenAI(api_key=k,**({"base_url":b} if b else {}))
    def chat_with_tools(self,m,t): r=self.c.chat.completions.create(model=self.m,messages=m,max_tokens=4096,tools=t,tool_choice="auto"); ch=r.choices[0]; msg={"role":"assistant","content":ch.message.content or ""}; 
        if ch.message.tool_calls: msg["tool_calls"]=[{"id":tc.id,"type":"function","function":{"name":tc.function.name,"arguments":tc.function.arguments}} for tc in ch.message.tool_calls]
        return {"message":msg}
    def chat_simple(self,s,u): return self.c.chat.completions.create(model=self.m,messages=[{"role":"system","content":s},{"role":"user","content":u}],max_tokens=4096).choices[0].message.content or ""

class OllamaProvider(AIProvider):
    def __init__(self,m="llama3",b="http://localhost:11434"): self.m=m; self.b=b
    def chat_with_tools(self,m,t): return {"message":{"role":"assistant","content":self.chat_simple("",m[-1].get('content',''))}}
    def chat_simple(self,s,u): return requests.post(f"{self.b}/api/chat",json={"model":self.m,"messages":[{"role":"system","content":s},{"role":"user","content":u}],"stream":False},timeout=300).json().get('message',{}).get('content','')

class DemoProvider(AIProvider):
    def chat_with_tools(self,m,t): return {"message":{"role":"assistant","content":self._d(m[-1].get('content','') if m else '')}}
    def chat_simple(self,s,u): return self._d(u)
    def _d(self,p):
        if "outline" in p.lower(): return json.dumps([{"slide_number":1,"title":"Cut Costs 40% with AI","layout":"title","key_points":["Board Mandate","Key Metric"]},{"slide_number":2,"title":"The Bottleneck","layout":"content","key_points":["60% manual processing","Hidden labor costs"]},{"slide_number":3,"title":"3 Automation Pillars","layout":"bullets","key_points":["IDP","Predictive Allocation","Autonomous Reports"]},{"slide_number":4,"title":"IDP vs Manual Data","layout":"two_column","key_points":["Time: 4h -> 12min","Error rate: 35% -> 2%"]},{"slide_number":5,"title":"Roadmap","layout":"bullets","key_points":["Q3 Pilot","Q4 Scale","Q1 Optimize"]},{"slide_number":6,"title":"Projected ROI","layout":"bullets","key_points":["$4.2M Saved","150% ROI in 18 Months"]},{"slide_number":7,"title":"Next Steps","layout":"content","key_points":["Approve budget","Appoint lead"]},{"slide_number":8,"title":"Thank You","layout":"title","key_points":["Questions?"]}])
        elif "content" in p.lower(): return json.dumps([{"number":1,"title":"Cut Costs 40% with AI","subtitle":"Strategic Mandate","body":"The board has mandated a 40% reduction in operational overhead by Q4.","bullets":[],"highlight":"\"We must digitize or die.\" — CEO","visual_suggestion":"Full bleed dark background, large white text '40%'"},{"number":2,"title":"The Bottleneck","body":"Analysts spend 60% of time on manual data entry.","bullets":["60% time manual entry","35% error rate in logs","$2.1M hidden labor costs"],"visual_suggestion":"Icon of a stopwatch, three large red/yellow/green metrics"},{"number":3,"title":"3 Pillars","bullets":["IDP: Extract data from PDFs instantly","Predictive Allocation: 3 weeks ahead","Autonomous Reports for the board"],"visual_suggestion":"Three columns with icons"},{"number":4,"title":"IDP vs Manual Data","body":"Pilot showed massive efficiency gains.","bullets":["Processing: 4h -> 12min (95% faster)","Error rate: 35% -> 2%","Capacity: +40%"],"visual_suggestion":"Stacked bar chart Before vs After"},{"number":5,"title":"Roadmap","bullets":["Q3: Pilot in Finance & Ops","Q4: Scale to all depts","Q1: Optimize & Handover"],"visual_suggestion":"Horizontal timeline, 3 distinct blocks"},{"number":6,"title":"Projected ROI","bullets":["Investment: $1.4M over 12 months","Annual Savings: $4.2M","Payback Period: 4 Months"],"highlight":"150% ROI in 18 months.","visual_suggestion":"Large '150%' center stage, small donut chart"},{"number":7,"title":"Next Steps","body":"To hit Q4 targets, we need immediate alignment.","bullets":["Approve $350k Q3 pilot","Appointcross-functional lead","Bi-weekly steering reviews"],"visual_suggestion":"Clean layout, three checkboxes, calendar icon"},{"number":8,"title":"Thank You","subtitle":"Let's build the future.","body":"","bullets":[],"highlight":"","visual_suggestion":"Minimalist closing slide."}])
        else: return json.dumps([{"number":i,"speaker_notes":["[Advance Slide]","Concept: Set the stage","Bridge: 'To understand why we are here..."],"full_script":"Good morning. Last quarter, the board asked: 'Why are we still doing things manually?' That question led here. I'll show you exactly how we cut costs by 40%.","timing_secs":35,"delivery_tip":"Lower pitch for gravity."},{"number":2,"speaker_notes":["Concept: Expose the pain point","Bridge": 'This is a structural financial bleed.'"],"full_script":"So where are we today? Analysts spend 60% of time just moving data. Error rate is 35%. Hidden labor: $2.1M/year. This isn't just an efficiency issue; it's a structural financial bleed.","timing_secs":40,"delivery_tip":"Pause after '$2.1 million'."}])

# ═════════════════════════════════════════════════════════════
# GENERATORS & WIDGETS (Unchanged from v2.3)
# ═══════════════════════════════════════════════════════════════
class PPTXGenerator:
    def __init__(self,t='executive_blue'): self.tpl=TEMPLATES.get(t,TEMPLATES['executive_blue'])
    def _rgb(self,t): return RGBColor(*t)
    def generate(self,slides,cfg,path):
        prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
        for sd in slides:
            s=prs.slides.add_slide(prs.slide_layouts[6])
            s.background.fill.solid(); s.background.fill.fore_color.rgb=self._rgb(self.tpl['bg'])
            sh=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(0),Inches(0),prs.slide_width,Inches(0.08)); sh.fill.solid(); sh.fill.fore_color.rgb=self._rgb(self.tpl['accent']); sh.line.fill.background()
            tb=s.shapes.add_textbox(Inches(12.5),Inches(7.05),Inches(0.7),Inches(0.35)); p=tb.text_frame.paragraphs[0]; p.text=f"{sd.number}/{len(slides)}"; p.font.size=Pt(10); p.font.color.rgb=self._rgb(self.tpl['body_c']); p.alignment=PP_ALIGN.RIGHT
            if sd.layout=='title': self._title(s,sd)
            elif sd.layout=='bullets': self._bullets(s,sd)
            elif sd.layout=='two_column': self._twocol(s,sd)
            else: self._content(s,sd)
            if cfg.include_notes and (sd.speaker_notes or sd.full_script):
                txt="\n".join(f"• {n}" for n in sd.speaker_notes)
                if cfg.include_script: txt+=f"\n\nSCRIPT:\n{sd.full_script}"
                s.notes_slide.notes_text_frame.text=txt
        prs.save(path)
    def _ttl(self,s,t,top,sz=32): tb=s.shapes.add_textbox(Inches(0.9),top,Inches(11.5),Inches(1.2)); p=tb.text_frame.paragraphs[0]; p.text=t; p.font.size=Pt(sz); p.font.bold=True; p.font.color.rgb=self._rgb(self.tpl['title_c']); p.font.name=self.tpl['title_font']
    def _title(self,s,sd): self._ttl(s,sd.title,Inches(2.5),44)
    def _content(self,s,sd):
        self._ttl(s,sd.title,Inches(0.4))
        if sd.body:
            tb=s.shapes.add_textbox(Inches(0.9),Inches(1.6),Inches(11.5),Inches(5))
            for i,pt in enumerate(sd.body.split('\n')): p=tb.text_frame.paragraphs[0] if i==0 else tb.text_frame.add_paragraph(); p.text=pt; p.font.size=Pt(self.tpl['body_sz']); p.font.color.rgb=self._rgb(self.tpl['body_c']); p.space_after=Pt(8)
    def _bullets(self,s,sd):
        self._ttl(s,sd.title,Inches(0.4))
        tb=s.shapes.add_textbox(Inches(1.2),Inches(1.6),Inches(11),Inches(5.2))
        for i,b in enumerate(sd.bullets): p=tb.text_frame.paragraphs[0] if i==0 else tb.text_frame.add_paragraph(); p.text=f"▸  {b}"; p.font.size=Pt(self.tpl['bullet_sz']); p.font.color.rgb=self._rgb(self.tpl['body_c']); p.space_after=Pt(14)
    def _twocol(self,s,sd):
        self._ttl(s,sd.title,Inches(0.4))
        if sd.body:
            tb=s.shapes.add_textbox(Inches(0.9),Inches(1.6),Inches(5.5),Inches(5))
            for i,pt in enumerate(sd.body.split('\n')): p=tb.text_frame.paragraphs[0] if i==0 else tb.text_frame.add_paragraph(); p.text=pt; p.font.size=Pt(self.tpl['body_sz']); p.font.color.rgb=self._rgb(self.tpl['body_c'])
        if sd.bullets:
            tb=s.shapes.add_textbox(Inches(7),Inches(1.6),Inches(5.5),Inches(5))
            for i,b in enumerate(sd.bullets): p=tb.text_frame.paragraphs[0] if i==0 else tb.text_frame.add_paragraph(); p.text=f"▸  {b}"; p.font.size=Pt(self.tpl['bullet_sz']); p.font.color.rgb=self._rgb(self.tpl['body_c'])
        ln=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Inches(6.5),Inches(1.6),Inches(0.02),Inches(5)); ln.fill.solid(); ln.fill.fore_color.rgb=self._rgb(self.tpl['accent']); ln.line.fill.background()

class PDFGenerator:
    def __init__(self,t='executive_blue'): self.tpl=TEMPLATES.get(t,TEMPLATES['executive_blue'])
    def slides_pdf(self,slides,path):
        if not HAS_REPORTLAB: raise ImportError("pip install reportlab")
        w,h=landscape(LETTER); c=canvas.Canvas(path,pagesize=(w,h))
        for sd in slides: c.setFillColor(HexColor('#%02x%02x%02x'%self.tpl['bg'])); c.rect(0,0,w,h,fill=1,stroke=0)
        c.setFillColor(HexColor('#%02x%02x%02x'%self.tpl['title_c'])); c.setFont(self.tpl['title_font'],self.tpl['title_sz']-4); c.drawString(50,h-60,sd.title[:80])
        c.setFillColor(HexColor('#%02x%02x%02x'%self.tpl['body_c'])); c.setFont(self.tpl['body_font'],self.tpl['body_sz']-2); y=h-110
        for ln in (sd.body or "").split('\n'):
            for wr in textwrap.wrap(ln,90):
                if y>40: c.drawString(50,y,wr); y-=22
        c.showPage()
        c.save()
    def notes_pdf(self,slides,path,script=True):
        if not HAS_REPORTLAB: raise ImportError("pip install reportlab")
        w,h=LETTER; c=canvas.Canvas(path,pagesize=(w,h)); c.setFillColor(white); c.rect(0,0,w,h,fill=1,stroke=0)
        for sd in slides: y=h-50; c.setFillColor(black); c.setFont("Helvetica-Bold",16); c.drawString(40,y,f"Slide {sd.number}: {sd.title}"); y-=35; c.setFont("Helvetica",11)
        for n in sd.speaker_notes:
            for wr in textwrap.wrap(f"• {n}",80):
                if y>50: c.drawString(50,y,wr); y-=16
        if script and sd.full_script: y-=20; c.setFont("Helvetica-Bold",12); c.drawString(40,y,"Script:"); y-=20; c.setFont("Helvetica",10)
        for ln in sd.full_script.split('\n'):
            for wr in textwrap.wrap(ln,85):
                if y>50: c.drawString(50,y,wr); y-=14
        c.showPage()
        c.save()

class MarkdownExporter:
    def generate(self, slides: List[SlideData], cfg: PresConfig) -> str:
        md = f"# {cfg.topic}\n\n**Audience:** {cfg.audience} | **Tone:** {cfg.tone}\n\n---\n\n"
        for s in slides:
            md += f"## Slide {s.number}: {s.title}\n"
            if s.body: md += f"{s.body}\n\n"
            if s.bullets: md += "\n".join(f"- {b}" for b in s.bullets) + "\n\n"
            if s.visual_suggestion: md += f"*Visual: {s.visual_suggestion}*\n\n"
            if s.full_script: md += f"### 📝 Script\n{s.full_script}\n\n"
            if s.metadata.get('sources'): md += f"**Sources**: {', '.join(s.metadata['sources'][:5])}\n\n"
            md += "---\n\n"
        return md

# ═════════════════════════════════════════════════════════════════════════════
# GUI WIDGETS
# ═══════════════════════════════════════════════════════════════════════════
class AccentBtn(QPushButton):
    def __init__(self,t,c=None,p=None): super().__init__(t,p); self.ac=c or C['blue']; self.setCursor(QCursor(Qt.PointingHandCursor); self.setMinimumHeight(42); self.setFont(QFont("Segoe UI",10,QFont.Bold))
    def _s(self): self.setStyleSheet(f"QPushButton{{background:{self.ac};color:{C['crust']};border:none;border-radius:8px;padding:8px 24px;font-weight:bold;}}QPushButton:hover{{background:{C['lavender']};}}")
class GhostBtn(QPushButton):
    def __init__(self,t,p=None): super().__init__(t,p); self.setCursor(QCursor(Qt.PointingHandCursor); self.setMinimumHeight(36); self.setFont(QFont("Segoe UI",9))
    def _s(): self.setStyleSheet(f"QPushButton{{background:transparent;color:{C['blue']};border:1px solid {C['surface1']};border-radius:6px;padding:6px 16px;}}QPushButton:hover{{background:{C['surface0']};}}")
class Card(QFrame):
    def __init__(self,p=None): super().__init__(p); self.setStyleSheet(f"Card{{background:{C['surface0']};border:1px solid {C['surface1']};border-radius:10px;padding:12px;}}")
class Steps(QWidget):
    clicked=Signal(int)
    def __init__(self,steps,p=None): super().__init__(p); self.steps=steps; self.cur=0; self.done=set(); self.setFixedWidth(220); self._b()
    def _b(self):
        l=QVBoxLayout(self); l.setContentsMargins(10,20,10,20); l.setSpacing(0)
        for i,n in enumerate(self.steps):
            b=QPushButton(f"  {i+1}. {n}"); b.setCursor(QCursor(Qt.PointingHandCursor); b.setObjectName(f"s{i}"); b.clicked.connect(lambda _,x=i: self.clicked.emit(x)); b.setFont(QFont("Segoe UI",10))
            b.setStyleSheet(f"QPushButton{{background:transparent;color:{C['overlay0']};border:none;border-radius:6px;padding:8px 10px;text-align:left;font-size:12px;}}QPushButton:hover{{background:{C['surface0']};}}")
            l.addWidget(b)
            if i<len(self.steps)-1: ln=QFrame(); ln.setFixedHeight(20); ln.setObjectName(f"l{i}"); ln.setStyleSheet(f"background:{C['surface1']};"); l.addWidget(ln,alignment=Qt.AlignHCenter)
        l.addStretch()
    def set_current(self,idx):
        self.cur=idx
        for i in range(len(self.steps)):
            b=self.findChild(QPushButton,f"s{i}"); ln=self.findChild(QFrame,f"l{i}")
            if b:
                if i==idx: b.setStyleSheet(f"QPushButton{{background:{C['surface0']};color:{C['blue']};border-left:3px solid {C['blue']};border-radius:0 6px 6px 0;padding:8px 10px;text-align:left;font-size:12px;font-weight:bold;}}")
                elif i in self.done: b.setStyleSheet(f"QPushButton{{background:transparent;color:{C['green']};border:none;border-radius:6px;padding:8px 10px;text-align:left;font-size:12px;}}")
                else: b.setStyleSheet(f"QPushButton{{background:transparent;color:{C['overlay0']};border:none;border-radius:6px;padding:8px 10px;text-align:left;font-size:12px;}}")
            if ln: ln.setStyleSheet(f"background:{C['green'] if i<idx or i in self.done else C['surface1']};")
class Overlay(QWidget):
    def __init__(self,p=None): super().__init__(p); self.setAttribute(Qt.WA_TransparentForMouseEvents,False); self.setVisible(False); self._d=0; self._t=QTimer(); self._t.timeout.connect(self._a); self._b=""
    def show_loading(self,t="Working"): self._b=t; self._d=0; self.setVisible(True); self.raise_(); self._t.start(400); self.update()
    def hide_loading(self): self._t.stop(); self.setVisible(False)
    def _a(self): self._d=(self._d+1)%4; self.update()
    def paintEvent(self,e):
        if not self.isVisible(): return
        p=QPainter(self); p.fillRect(self.rect(),QColor(0,0,0,160)); p.setPen(QColor(*[int(C['text'][i:i+2],16) for i in (1,3,5)])); p.setFont(QFont("Segoe UI",16,QFont.Bold)); p.drawText(self.rect(),Qt.AlignCenter,f"{self._b}{'.'*self._d}"); p.end()
    def resizeEvent(self,e):
        if self.parentWidget(): self.setGeometry(self.parentWidget().rect())
class SlideCard(Card):
    edit_req=Signal(int); up_req=Signal(int); down_req=Signal(int)
    def __init__(self,s,i,is_f,is_l,p=None):
        super().__init__(p); lay=QVBoxLayout(self); lay.setSpacing(6); lay.setContentsMargins(12,12,12,12); h=QHBoxLayout()
        up=QPushButton("▲"); up.setFixedSize(24,24); up.setCursor(QCursor(Qt.PointingHandCursor); up.setEnabled(not is_f); up.clicked.connect(lambda: self.up_req.emit(i)); up.setStyleSheet(f"QPushButton{{background:{C['surface1']};color:{C['text']};border:none;border-radius:4px;}}QPushButton:disabled{{color:{C['overlay0'];}}")
        dn=QPushButton("▼"); dn.setFixedSize(24,24); dn.setCursor(QCursor(Qt.PointingHandCursor); dn.setEnabled(not is_l); dn.clicked.connect(lambda: self.down_req.emit(i)); dn.setStyleSheet(up.styleSheet())
        nl=QLabel(f"#{s.number}"); nl.setStyleSheet(f"color:{C['blue']};font-weight:bold;font-size:14px;")
        tl=QLabel(s.title); tl.setStyleSheet(f"color:{C['text']};font-size:14px;font-weight:bold;"); tl.setWordWrap(True)
        ll=QLabel(s.layout.upper()); ll.setStyleSheet(f"color:{C['peach']};font-size:9px;font-weight:bold;background:{C['surface1']};border-radius:4px;padding:2px 8px;")
        h.addWidget(up); h.addWidget(dn); h.addWidget(nl); h.addWidget(tl,1); h.addWidget(ll); lay.addLayout(h)
        for bp in s.bullets[:2]: bl=QLabel(f"  • {bp[:60]}..."); bl.setStyleSheet(f"color:{C['subtext']};font-size:11px;"); bl.setWordWrap(True); lay.addWidget(bl)
        eb=GhostBtn("✏️ Edit"); eb.setFixedWidth(70); eb.clicked.connect(lambda: self.edit_req.emit(i)); lay.addWidget(eb,alignment=Qt.AlignRight)
class ToolLogPanel(QWidget):
    def __init__(self, p=None):
        super().__init__(p); lay=QVBoxLayout(self); lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)
        hdr=QHBoxLayout(); hdr.addWidget(QLabel("📡 Agent Log")); hdr.addStretch()
        clr=GhostBtn("Clear"); clr.setFixedHeight(24); clr.clicked.connect(lambda: self.log.clear()); hdr.addWidget(clr); lay.addLayout(hdr)
        self.log=QTextEdit(); self.log.setReadOnly(True); self.log.setMaximumHeight(150)
        self.log.setStyleSheet(f"background:{C['crust']};color:{C['subtext']};border:1px solid {C['surface0']};border-radius:6px;font-family:Consolas,monospace;font-size:11px;padding:8px;")
        lay.addWidget(self.log)
    def append_log(self, msg):
        c=C['green'] if msg.startswith("✅") else C['yellow'] if msg.startswith("🔧") else C['red'] if msg.startswith("❌") else C['text']
        self.log.append(f"<span style='color:{c};'>{msg}</span>"); self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())
class Toast(QFrame):
    def __init__(self, msg, color, parent=None): super().__init__(parent); self.setFixedSize(320,50); self.setStyleSheet(f"background:{color};color:{C['crust']};border-radius:10px;padding:12px;font-weight:bold;font-size:13px;"); lay=QHBoxLayout(self); lay.addWidget(QLabel(msg)); self._parent=parent
    def show_toast(self): 
        if self._parent: geo=self._parent.geometry(); self.move(geo.right()-self.width()-20,geo.bottom()-self.height()-40)
        self.show(); self.raise_(); QTimer.singleShot(3000, self.close)
def _lbl(t): l=QLabel(t); l.setStyleSheet(f"color:{C['text']};font-weight:bold;font-size:12px;"); l.setFixedWidth(120); return l
def _inp(): return f"background:{C['surface0']};color:{C['text']};border:1px solid {C['surface1']};border-radius:6px;padding:8px;font-size:12px;"

# ═════════════════════════════════════════════════════════════════
# WORKFLOW PAGES
# ═════════════════════════════════════════════════════════════════════════════════
class ConfigPage(QWidget):
    ready=Signal()
    def __init__(self,p=None): super().__init__(p); self.gw=GPUWidget(); self.gw.setVisible(False); self._b()
    def _b(self):
        m=QVBoxLayout(self); m.setContentsMargins(40,30,40,30); t=QLabel("⚙️  AI Configuration"); t.setFont(QFont("Segoe UI",22,QFont.Bold)); m.addWidget(t); m.addSpacing(20)
        c=Card(); g=QGridLayout(c); g.setSpacing(12)
        g.addWidget(_lbl("Provider"),0,0); self.pc=QComboBox(); self.pc.addItems(["Demo Mode (No Key)","OpenAI","Ollama (Local)","🦙 Llama.cpp (Production)"]); self.pc.currentIndexChanged.connect(self._op); self.pc.setMinimumWidth(300); g.addWidget(self.pc,0,1,1,2)
        g.addWidget(_lbl("API Key"),1,0); self.ak=QLineEdit(); self.ak.setPlaceholderText("sk-..."); self.ak.setEchoMode(QLineEdit.Password); g.addWidget(self.ak,1,1,1,2)
        g.addWidget(_lbl("Model"),2,0); self.mc=QComboBox(); self.mc.addItems(["demo"]); self.mc.setEditable(True); g.addWidget(self.mc,2,1,1,2)
        m.addWidget(c); self.gw=GPUWidget(); self.gw.setVisible(False); m.addWidget(self.gw); m.addStretch(); r=QHBoxLayout(); r.addStretch(); self.nb=AccentBtn("Continue →"); self.nb.clicked.connect(self.ready); r.addWidget(self.nb); m.addLayout(r); self._op(0)
    def _op(self,i):
        d,o,l=i==0,i==2,i==3; self.ak.setEnabled(not d and not o and not l); self.gw.setVisible(l); self.mc.clear()
        if d: self.mc.addItems(["demo"])
        elif o: self.mc.addItems(["llama3","llama3.1","mistral"]); self.mc.setEditable(True)
        elif l: self.mc.addItems(["(select GGUF below)"]); self.mc.setEnabled(False)
        else: self.mc.addItems(["gpt-4o-mini","gpt-4o"]); self.mc.setEditable(True)
    def get_provider(self):
        i=self.pc.currentIndex()
        if i==0: return DemoProvider()
        elif i==2: return OllamaProvider(self.mc.currentText())
        elif i==3: self.gw.gpu.model_path=self.gw.mp.currentText(); return LlamaCppProvider(self.gw.gpu)
        else: k=self.ak.text().strip()
        if not k: raise ValueError("Enter API key")
        return OpenAIProvider(k,self.mc.currentText())
    def get_gpu(self): return self.gw.gpu if self.pc.currentIndex()==3 else None
    def get_tools(self,cfg): return PresentationTools(cfg,self.get_gpu())

class BriefPage(QWidget):
    go=Signal()
    def __init__(self,p=None): super().__init__(p); self._b()
    def _b(self):
        m=QVBoxLayout(self); m.setContentsMargins(40,30,40,30); t=QLabel("📝  Brief"); t.setFont(QFont("Segoe UI",22,QFont.Bold)); m.addWidget(t); m.addSpacing(20)
        c=Card(); f=QGridLayout(c); f.setSpacing(14)
        f.addWidget(_lbl("Topic *"),0,0); self.tp=QLineEdit(); self.tp.setPlaceholderText("e.g., Cut Operational Costs 40% with AI"); self.tp.setMinimumHeight(40); f.addWidget(self.tp,0,1)
        f.addWidget(_lbl("Audience"),1,0); self.au=QComboBox(); self.au.addItems(["Executive leadership","Technical team","Investors / Board","General professionals","Sales & marketing"]); self.au.setEditable(True); f.addWidget(self.au,1,1)
        f.addWidget(_lbl("Slides"),2,0); self.ns=QSpinBox(); self.ns.setRange(5,30); self.ns.setValue(8); f.addWidget(self.ns,2,1)
        f.addWidget(_lbl("Tone"),3,0); self.tn=QComboBox(); self.tn.addItems(["Executive","Persuasive","Technical","Academic","Inspirational"]); f.addWidget(self.tn,3,1)
        f.addWidget(_lbl("Extra"),4,0,Qt.AlignTop); self.ex=QTextEdit(); self.ex.setPlaceholderText("Specific data points to include..."); self.ex.setMaximumHeight(60); f.addWidget(self.ex,4,1)
        m.addWidget(c)
        self.pro_toggle=QPushButton("⚙️ Pro Mode"); self.pro_toggle.setCheckable(True); self.pro_toggle.setCursor(QCursor(Qt.PointingHandCursor); self.pro_toggle.setStyleSheet(f"QPushButton{{background:transparent;color:{C['mauve'];border:none;font-size:11px;}}")
        self.pro_toggle.toggled.connect(lambda c: self.pro_card.setVisible(c)); m.addWidget(self.pro_toggle)
        self.pro_card=Card(); self.pro_card.setVisible(False); pl=QVBoxLayout(self.pro_card)
        self.sys_ov=QTextEdit(); self.sys_ov.setMaximumHeight(80); self.sys_ov.setPlaceholderText("You are an expert..."); self.sys_ov.setStyleSheet(_inp()); pl.addWidget(self.sys_ov); m.addWidget(self.pro_card)
        m.addStretch(); r=QHBoxLayout(); r.addStretch(); self.gb=AccentBtn("🚀 Generate Outline"); self.gb.clicked.connect(self.go); r.addWidget(self.gb); m.addLayout(r)
    def get_config(self): return PresConfig(self.tp.text().strip(),self.au.currentText(),self.ns.value(),self.tn.currentText(),extra=self.ex.toPlainText().strip(),system_override=self.sys_ov.toPlainText().strip())

class OutlinePage(QWidget):
    approve=Signal(); back=Signal()
    def __init__(self,p=None): super().__init__(p); self.slides=[]; self._b()
    def _b(self):
        m=QVBoxLayout(self); m.setContentsMargins(40,30,40,30); h=QHBoxLayout(); t=QLabel("📋  Outline"); t.setFont(QFont("Segoe UI",22,QFont.Bold)); h.addWidget(t); self.rg=GhostBtn("🔄 Regen"); h.addWidget(self.rg); m.addLayout(h)
        self.sb=QLabel(""); self.sb.setStyleSheet(f"color:{C['subtext']};font-size:13px;"); m.addWidget(self.sb); m.addSpacing(10)
        sc=QScrollArea(); sc.setWidgetResizable(True); sc.setStyleSheet(f"QScrollArea{{border:none;background:transparent;}}QScrollBar:vertical{{background:{C['base']};width:8px;border-radius:4px;}}QScrollBar::handle:vertical{{background:{C['surface1']};border-radius:4px;min-height:30px;}}")
        self.ct=QWidget(); self.cl=QVBoxLayout(self.ct); self.cl.setSpacing(10); self.cl.addStretch(); sc.setWidget(self.ct); m.addWidget(sc,1)
        r=QHBoxLayout(); self.bb=GhostBtn("← Back"); self.bb.clicked.connect(self.back); r.addWidget(self.bb); r.addStretch()
        self.ab=AccentBtn("✅ Approve →"); self.ab.clicked.connect(self.approve); r.addWidget(self.ab); m.addLayout(r)
    def set_slides(self,sl):
        self.slides=sl
        while self.cl.count(): it=self.cl.takeAt(0); (it.widget().deleteLater() if it.widget() else None)
        for i,s in enumerate(sl):
            cd=SlideCard(s,i,i==0,i==len(sl)-1); cd.edit_req.connect(self._ed); cd.up_req.connect(self._up); cd.down_req.connect(self._dn)
            self.cl.insertWidget(self.cl.count()-1,cd)
        self.sb.setText(f"✅ {len(sl)} slides. ▲▼ to reorder.")
    def _ed(self,i):
        if 0<=i<len(self.slides):
            s=self.slides[i]; d=QMessageBox(self); d.setWindowTitle(f"Edit #{s.number}"); d.setStyleSheet(f"QMessageBox{{background:{C['base']};}}QLabel{{color:{C['text']};}}")
            lo=d.layout(); te=QLineEdit(s.title); te.setStyleSheet(_inp()); lo.addWidget(QLabel("Title:"),1,0); lo.addWidget(te,1,1,1,2)
            lc=QComboBox(); lc.addItems(["title","content","bullets","two_column"]); lc.setCurrentText(s.layout); lc.setStyleSheet(_inp()); lo.addWidget(QLabel("Layout:"),2,0); lo.addWidget(lc,2,1,1,2)
            be=QTextEdit("\n".join(s.bullets)); be.setMaximumHeight(120); be.setStyleSheet(_inp()); lo.addWidget(QLabel("Points:"),3,0,1,3); lo.addWidget(be,4,0,1,3)
            d.addButton("Cancel",QMessageBox.RejectRole); d.addButton("Save",QMessageBox.AcceptRole)
            if d.exec()==QMessageBox.AcceptRole: s.title=te.text(); s.layout=lc.currentText(); s.bullets=[l.strip() for l in be.toPlainText().split('\n') if l.strip()]; self.set_slides(self.slides)
    def _up(self,i):
        if i>0: self.slides[i-1], self.slides[i] = self.slides[i], self.slides[i-1]; self.set_slides(self.slides)
    def _dn(self,i):
        if i<len(self.slides)-1: self.slides[i], self.slides[i+1] = self.slides[i+1], self.slides[i]; self.set_slides(self.slides)

class ContentPage(QWidget):
    go=Signal(); back=Signal()
    def __init__(self,p=None): super().__init__(p); self.slides=[]; self._b()
    def _b(self):
        m=QVBoxLayout(self); m.setContentsMargins(40,20,40,20); h=QHBoxLayout()
        t=QLabel("📄  Content"); t.setFont(QFont("Segoe UI",22,QFont.Bold)); h.addWidget(t); self.rg=GhostBtn("🔄 Regen"); h.addWidget(self.rg); self.wc_label=QLabel("0 words"); self.wc_label.setStyleSheet(f"color:{C['overlay0']};font-size:11px;"); h.addWidget(self.wc_label)
        m.addLayout(h)
        self.tabs=QTabWidget(); self.tabs.setStyleSheet(f"QTabWidget::pane{{border:1px solid {C['surface1']};background:{C['base']};border-radius:8px;}}QTabBar::tab{{background:{C['surface0']};color:{C['subtext']};padding:8px 16px;border-top-left-radius:6px;border-top-right-radius:6px;margin-right:2px;font-size:11px;}}QTabBar::tab:selected{{background:{C['surface1']};color:{C['blue']};font-weight:bold;}}")
        self.tabs.currentChanged.connect(self._update_wc); m.addWidget(self.tabs,1)
        r=QHBoxLayout(); self.bb=GhostBtn("← Back"); self.bb.clicked.connect(self.back); r.addWidget(self.bb); r.addStretch()
        self.nb=AccentBtn("🎤 Generate Notes →"); self.nb.clicked.connect(self.go); r.addWidget(self.nb); m.addLayout(r)
    def _update_wc(self):
        idx=self.tabs.currentIndex()
        if 0<=idx<len(self.slides): self.wc_label.setText(f"Slide {idx+1}: {len((self.slides[idx].body+' '+' '.join(self.slides[idx].bullets)).split())} words")
    def set_slides(self,sl):
        self.slides=sl; self.tabs.clear()
        for i,s in enumerate(sl):
            p=QWidget(); lay=QVBoxLayout(p); lay.setContentsMargins(16,16,16,16); lay.setSpacing(10)
            for lt,wf in [("Title",lambda _i=i:QLineEdit(self.slides[_i].title)),("Body",lambda _i=i:self._te(self.slides[_i].body,120)),("Bullets",lambda _i=i:self._te("\n".join(self.slides[_i].bullets),100)),("Highlight",lambda _i=i:QLineEdit(self.slides[_i].highlight)),("Visual",lambda _i=i:QLineEdit(self.slides[_i].visual_suggestion))]:
                lb=QLabel(lt); lb.setStyleSheet(f"color:{C['text']};font-weight:bold;font-size:11px;"); lay.addWidget(lb); w=wf(); w.setStyleSheet(_inp()); lay.addWidget(w)
            lay.addStretch(); self.tabs.addTab(p,f"#{s.number}")
    def _te(self,t,h): e=QTextEdit(t); e.setMinimumHeight(h); e.setStyleSheet(_inp()); return e

class NotesPage(QWidget):
    go=Signal(); back=Signal()
    def __init__(self,p=None): super().__init__(p); self.slides=[]; self._b()
    def _b(self):
        m=QVBoxLayout(self); m.setContentsMargins(40,20,40,20); h=QHBoxLayout()
        t=QLabel("🎤  Notes"); t.setFont(QFont("Segoe UI",22,QFont.Bold)); h.addWidget(t); self.rg=GhostBtn("🔄 Regen"); h.addWidget(self.rg)
        self.wc_label=QLabel("0 words"); self.wc_label.setStyleSheet(f"color:{C['overlay0']};font-size:11px;"); h.addWidget(self.wc_label)
        m.addLayout(h)
        self.tabs=QTabWidget(); self.tabs.setStyleSheet(f"QTabWidget::pane{{border:1px solid {C['surface1']};background:{C['base']};border-radius:8px;}}QTabBar::tab{{background:{C['surface0']};color:{C['subtext']};padding:8px 16px;border-top-left-radius:6px;border-top-right-radius:6px;margin-right:2px;font-size:11px;}}QTabBar::tab:selected{{background:{C['surface1']};color:{C['mauve']};font-weight:bold;}}")
        self.tabs.currentChanged.connect(self._update_wc); m.addWidget(self.tabs,1)
        r=QHBoxLayout(); self.bb=GhostBtn("← Back"); self.bb.clicked.connect(self.back); r.addWidget(self.bb); r.addStretch()
        self.eb=AccentBtn("📦 Export →",C['mauve']); self.eb.clicked.connect(self.go); r.addWidget(self.eb); m.addLayout(r)
    def _update_wc(self):
        idx=self.tabs.currentIndex()
        if 0<=idx<len(self.slides): self.wc_label.setText(f"Script {idx+1}: {len(self.slides[idx].full_script.split())} words")
    def set_slides(self,sl):
        self.slides=sl; self.tabs.clear()
        for i,s in enumerate(sl):
            p=QWidget(); lay=QVBoxLayout(p); lay.setContentsMargins(16,16,16,16); lay.setSpacing(10)
            inf=QLabel(f"#{s.number}: {s.title}  ⏱{s.timing}s  💡{s.tips}"); inf.setStyleSheet(f"color:{C['teal']};font-size:11px;padding:6px;background:{C['surface0']};border-radius:4px;"); inf.setWordWrap(True); lay.addWidget(inf)
            lb=QLabel("Speaker Notes:"); lb.setStyleSheet(f"color:{C['text']};font-weight:bold;font-size:11px;"); lay.addWidget(lb)
            ne=QTextEdit("\n".join(f"• {n}" for n in s.speaker_notes)); ne.setMinimumHeight(80); ne.setStyleSheet(_inp()); lay.addWidget(ne)
            cp_row=QHBoxLayout(); lb2=QLabel("Full Script:"); lb2.setStyleSheet(f"color:{C['text']};font-weight:bold;font-size:11px;"); cp_row.addWidget(lb2); cp_row.addStretch()
            cp_btn=GhostBtn("📋 Copy Script")
            cp_btn.clicked.connect(lambda checked, txt=s.full_script: (QApplication.clipboard().setText(txt), self.parent().parent().parent().show_toast("Copied!", C['green'])))
            cp_row.addWidget(cp_btn); lay.addLayout(cp_row)
            se=QTextEdit(s.full_script); se.setMinimumHeight(150); se.setStyleSheet(_inp()); lay.addWidget(se)
            lay.addStretch(); self.tabs.addTab(p,f"#{s.number}")

class ExportPage(QWidget):
    back=Signal()
    def __init__(self,p=None): super().__init__(p); self.slides=[]; self.cfg=PresConfig(); self._b()
    def _b(self):
        m=QVBoxLayout(self); m.setContentsMargins(40,30,40,30); t=QLabel("📦  Export & Stats"); t.setFont(QFont("Segoe UI",22,QFont.Bold)); m.addWidget(t); m.addSpacing(15)
        self.stats_card=Card(); sl=QHBoxLayout(self.stats_card); self.stat_labels={}
        for k,ic in [("slides","📊"),("words","📝"),("time","⏱️")]:
            w=QWidget(); wl=QVBoxLayout(w); wl.setSpacing(2)
            l1=QLabel(ic); l1.setAlignment(Qt.AlignCenter); l1.setFont(QFont("Segoe UI",16))
            v=QLabel("0"); v.setAlignment(Qt.AlignCenter); v.setStyleSheet(f"color:{C['blue']};font-weight:bold;font-size:18px;")
            l2=QLabel(k.capitalize()); l2.setAlignment(Qt.AlignCenter); l2.setStyleSheet(f"color:{C['overlay0']};font-size:10px;")
            wl.addWidget(l1); wl.addWidget(v); wl.addWidget(l2); sl.addWidget(w); self.stat_labels[k]=v
        m.addWidget(self.stats_card); m.addSpacing(10)
        tc=Card(); tl=QVBoxLayout(tc); tl.addWidget(_lbl("Template"))
        tr=QHBoxLayout(); self.tb={}
        for k,tp in TEMPLATES.items():
            b=QPushButton(f"{tp['icon']}  {tp['name']}"); b.setCheckable(True); b.setCursor(QCursor(Qt.PointingHandCursor); b.setFixedHeight(50); b.setMinimumWidth(160)
            b.setStyleSheet(f"QPushButton{{background:{C['surface1']};color:{C['text']};border:2px solid transparent;border-radius:8px;font-size:12px;font-weight:bold;}}QPushButton:checked{{border-color:{C['blue']};background:{C['surface0']};color:{C['blue']};}}")
            b.clicked.connect(lambda _,x=k: self._st(x)); tr.addWidget(b); self.tb[k]=b
        self.tb['executive_blue'].setChecked(True); tl.addLayout(tr); m.addWidget(tc)
        oc=Card(); ol=QVBoxLayout(oc)
        self.nc=QCheckBox("Include notes"); self.nc.setChecked(True); self.nc.setStyleSheet(f"color:{C['text']};font-size:12px;"); ol.addWidget(self.nc)
        self.sc=QCheckBox("Include script"); self.sc.setChecked(True); self.sc.setStyleSheet(f"color:{C['text']};font-size:12px;"); ol.addWidget(self.sc)
        m.addWidget(oc); m.addSpacing(10)
        ec=Card(); el=QVBoxLayout(ec)
        b1=AccentBtn("📋  PowerPoint (.pptx"); b1.setMinimumHeight(44); b1.clicked.connect(self._xp); el.addWidget(b1)
        b5=AccentBtn("📝  Markdown (.md)", C['teal']); b5.setMinimumHeight(44); b5.clicked.connect(self._xmd); el.addWidget(b5)
        pr=QHBoxLayout()
        b2=AccentBtn("📄 PDF Slides",C['peach']); b2.setMinimumHeight(44); b2.clicked.connect(self._xs); pr.addWidget(b2)
        b3=AccentBtn("🎤 PDF Notes",C['mauve']); b3.setMinimumHeight(44); b3.clicked.connect(self._xn); pr.addWidget(b3)
        m.addWidget(ec)
        self.st=QLabel(""); self.st.setStyleSheet(f"color:{C['subtext']};font-size:11px;"); self.st.setWordWrap(True); m.addWidget(self.st)
        m.addStretch(); r=QHBoxLayout(); self.bb=GhostBtn("← Back"); self.bb.clicked.connect(self.back); r.addWidget(self.bb); r.addStretch(); m.addLayout(r)
    def update_stats(self,slides):
        self.slides=slides; self.stat_labels["slides"].setText(str(len(slides)))
        w=sum(len((s.body+" "+s.full_script+" "+" ".join(s.bullets)).split()) for s in slides)
        self.stat_labels["words"].setText(f"{w:,}"); t=sum(s.timing for s in slides); m2,s2=divmod(t,60); self.stat_labels["time"].setText(f"{m2}m {s2}s")
    def _st(self,k):
        for x,b in self.tb.items(): b.setChecked(x==k)
    def _gt(self): return next((k for k,b in self.tb.items() if b.isChecked()),'executive_blue')
    def _ss(self,mg,c=None): self.st.setText(mg); self.st.setStyleSheet(f"color:{c or C['subtext']};font-size:11px;")
    def _sp(self,f): p,_=QFileDialog.getSaveFileName(self,"Save","",f); return p
    def _of(self,p):
        if QMessageBox.question(self,"Open?",f"Saved:\n{p}\nOpen?",QMessageBox.Yes|QMessageBox.No)==QMessageBox.Yes:
            import subprocess
            if platform.system()=='Windows': os.startfile(p)
            elif platform.system()=='Darwin': subprocess.run(['open',p])
            else: subprocess.run(['xdg-open',p])
    def _xp(self):
        p=self._sp("PPTX (*.pptx)"); if not p: return
        if not p.endswith('.pptx'): p+='.pptx'
        self.cfg.template=self._gt(); self.cfg.include_notes=self.nc.isChecked(); self.cfg.include_script=self.sc.isChecked()
        try: PPTXGenerator(self.cfg.template).generate(self.slides,self.cfg,p); self._ss(f"✅ {p}",C['green']); self.parent().parent().parent().show_toast("PPTX Saved!", C['green']); self._of(p)
        except Exception as e: self._ss(f"❌ {e}",C['red'])
    def _xs(self):
        if not HAS_REPORTLAB: QMessageBox.warning(self,"Missing","pip install reportlab"); return
        p=self._sp("PDF (*.pdf)"); if not p: return
        if not p.endswith('.pdf'): p+='.pdf'
        try: PDFGenerator(self._gt()).slides_pdf(self.slides,p); self._ss(f"✅ {p}",C['green']); self.parent().parent().parent().show_toast("PDF Saved!", C['green']); self._of(p)
        except Exception as e: self._ss(f"❌ {e}",C['red'])
    def _xn(self):
        if not HAS_REPORTLAB: QMessageBox.warning(self,"Missing","pip install reportlab"); return
        p=self._sp("PDF (*.pdf)"); if not p: return
        if not p.endswith('.pdf'): p+='.pdf'
        try: PDFGenerator(self._gt()).notes_pdf(self.slides,p); self._ss(f"✅ {p}",C['green']); self.parent().parent().parent().show_toast("Notes PDF Saved!", C['green']); self._of(p)
        except Exception as e: self._ss(f"❌ {e}",C['red'])
    def _xmd(self):
        p=self._sp("Markdown (*.md)"); if not p: return
        if not p.endswith('.md'): p+='.md'
        try: md=MarkdownExporter().generate(self.slides, self.cfg); open(p,'w',encoding='utf-8').write(md); self._ss(f"✅ {p}",C['green']); self.parent().parent().parent().show_toast("Markdown Saved!", C['green']); self._of(p)
        except Exception as e: self._ss(f"❌ {e}",C['red'])

# ═══════════════════════════════════════════════════════════════════════
# GPU WIDGET (Model Folder Manager)
# ═════════════════════════════════════════════════════════════════════════════
class GPUWidget(Card):
    SETTINGS_FILE = "presagent_settings.json"
    DEFAULT_FOLDERS = ["./models", "~/.ollama/models", "~/models"]
    changed=Signal()
    def __init__(self,p=None): super().__init__(p); self.gpu=GPUConfig(); self._b()
    def _b(self):
        lay=QVBoxLayout(self); lay.setSpacing(10)
        r=QHBoxLayout(); r.addWidget(self._l("Model:")); self.mp=QComboBox(); self.mp.setEditable(True); self.mp.setPlaceholderText("Select or type path..."); r.addWidget(self.mp,1)
        scan=GhostBtn("📂 Scan All"); scan.clicked.connect(self._scan); r.addWidget(scan); br=GhostBtn("Browse File"); br.clicked.connect(self._browse); r.addWidget(br); lay.addLayout(r)
        lay.addWidget(self._l("Scan Folders:"))
        self.folder_list=QListWidget(); self.folder_list.setMaximumHeight(90)
        self.folder_list.setStyleSheet(f"QListWidget{{background:{C['crust']};color:{C['subtext']};border:1px solid {C['surface0']};border-radius:4px;font-size:10px;padding:4px;outline:none;}}QListWidget::item{{padding:3px;border-radius:2px;}}QListWidget::item:hover{{background:{C['surface0']};color:{C['text']}}QListWidget::item:selected{{background:{C['surface1']};color:{C['text']}}}}")
        lay.addWidget(self.folder_list)
        fbtn_row=QHBoxLayout()
        add_f=GhostBtn("+ Add"); add_f.clicked.connect(self._add_folder); fbtn_row.addWidget(add_f)
        rm_f=GhostBtn("- Remove"); rm_f.clicked.connect(self._rm_folder); fbtn_row.addWidget(rm_f)
        rst_f=GhostBtn("↺ Reset"); rst_f.clicked.connect(self._rst_folders); fbtn_row.addWidget(rst_f); lay.addLayout(fbtn_row)
        lay.addWidget(self._l("GPU Layers:"))
        gr=QHBoxLayout(); self.gl=QLabel("CPU (0)"); self.gl.setFixedWidth(100); self.gl.setStyleSheet(f"color:{C['yellow']};font-size:11px;font-weight:bold;")
        self.gs=QSlider(Qt.Horizontal); self.gs.setRange(-1,99); self.gs.setValue(0); self.gs.valueChanged.connect(self._gsl); gr.addWidget(self.gl); gr.addWidget(self.gs,1); lay.addLayout(gr)
        pr=QHBoxLayout()
        self.sb=AccentBtn("▶ Start",C['green']); self.sb.clicked.connect(self._start); pr.addWidget(self.sb)
        self.stb=GhostBtn("⏹ Stop"); self.stb.clicked.connect(self._stop); pr.addWidget(self.stb)
        self.hb=GhostBtn("💓 Health"); self.hb.clicked.connect(self._health); pr.addWidget(self.hb); lay.addLayout(pr)
        self.ss=QLabel("Ready"); self.ss.setStyleSheet(f"color:{C['overlay0']};font-size:11px;"); self.ss.setWordWrap(True); lay.addWidget(self.ss)
        self._init_folders()
    def _l(self,t): l=QLabel(t); l.setStyleSheet(f"color:{C['text']};font-weight:bold;font-size:11px;"); return l
    def _gsl(self,v): self.gpu.n_gpu_layers=v; self.gl.setText({"-1":"Auto GPU","0":"CPU (0)"}.get(str(v),f"{v} layers"))
    def _load_settings():
        if os.path.exists(self.SETTINGS_FILE):
            try: return json.load(open(self.SETTINGS_FILE, 'r')).get("custom_folders", [])
            except: pass
        return []
    def _save_settings(self):
        custom=[self.folder_list.item(i).text() for i in range(self.folder_list.count()) if self.folder_list.item(i).text() not in self.DEFAULT_FOLDERS]
        open(self.SETTINGS_FILE, 'w').write(json.dumps({"custom_folders": custom}, indent=2))
    def _init_folders(self):
        self.folder_list.clear()
        for f in self.DEFAULT_FOLDERS + self._load_settings(): self.folder_list.addItem(f)
    def _add_folder(self):
        p=QFileDialog.getExistingDirectory(self,"Select Model Folder")
        if p and not any(self.folder_list.item(i).text()==p for i in range(self.folder_list.count())):
            self.folder_list.addItem(p); self._save_settings()
    def _rm_folder(self):
        r=self.folder_list.currentRow()
        if r>=0: self.folder_list.takeItem(r); self._save_settings()
    def _rst_folders(self): self._init_folders(); self._save_settings(); self.ss.setText("ℹ️ Reset to defaults"); self.ss.setStyleSheet(f"color:{C['blue']};font-size:11px;")
    def _scan(self):
        self.mp.clear(); found=set()
        for i in range(self.folder_list.count()):
            p=Path(self.folder_list.item(i).text()).expanduser().resolve()
            if p.exists(): found.update(str(f) for f in p.rglob("*.gguf"))
        if found: self.mp.addItems(sorted(list(found), key=lambda x: os.path.basename(x)); self.ss.setText(f"✅ Found {len(found)} models."); self.ss.setStyleSheet(f"color:{C['green']};font-size:11px;")
        else: self.ss.setText("⚠️ No .gguf found. Add folders below."); self.ss.setStyleSheet(f"color:{C['yellow']};font-size:11px;")
    def _browse(self):
        p,_=QFileDialog.getOpenFileName(self,"GGUF","","GGUF (*.gguf);;All (*)")
        if p: self.mp.setCurrentText(p)
    def _start(self):
        self.gpu.model_path=self.mp.currentText(); self.ss.setText("⏳ Starting..."); self.ss.setStyleSheet(f"color:{C['yellow']};font-size:11px;"); QApplication.processEvents()
        ok,m=LlamaCppProvider(self.gpu).start_server(); self.ss.setText(f"{'✅' if ok else '❌'} {m}"); self.ss.setStyleSheet(f"color:{C['green'] if ok else C['red']};font-size:11px;")
    def _stop(self): LlamaCppProvider(self.gpu).stop_server(); self.ss.setText("Stopped")
    def _health(self):
        ok,m=LlamaCppProvider(self.gpu).health_check(); self.ss.setText(f"{'✅' if ok else '❌'} {m}"); self.ss.setStyleSheet(f"color:{C['green'] if ok else C['red']};font-size:11px;")

# ═════════════════════════════════════════════════════════════════════════════
# WORKER THREAD
# ═════════════════════════════════════════════════════════════════════════════════
class SkillWorker(QThread):
    progress=Signal(str); result=Signal(object); error=Signal(str); done=Signal()
    def __init__(self,prov,harness,skill,cfg,slides=None):
        super().__init__(); self.prov=prov; self.harness=harness; self.skill=skill; self.cfg=cfg; self.slides=slides or []
    def run(self):
        try:
            self.progress.emit(f"Running: {self.skill.name}...")
            if self.harness and not isinstance(self.prov,DemoProvider): r=self.harness.execute(self.skill,self.cfg,self.slides)
            else: time.sleep(0.5); r=self.skill.parse(self.prov.chat_simple(self.skill.sys(self.cfg),self.skill.usr(self.cfg,self.slides)),self.slides)
            self.result.emit(r)
        except Exception as e: self.error.emit(str(e))
        finally: self.done.emit()

# ═══════════════════════════════════════════════════════════════════════════════════════════════════
# MAIN WINDOW
# ═════════════════════════════════════════════════════════════════════════════════════════
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI PresAgent v2.5 — Web Search Integration")
        self.setMinimumSize(1100,750); self.resize(1350,850)
        self.prov=None; self.tools=None; self.harness=None; self.cfg=PresConfig()
        self.slides=[]; self.cur=0; self.project_path=None
        self.steps=["Configure AI","Brief","Outline","Content","Notes","Export"]
        self._theme(); self._ui(); self._connect()
    def _theme(self):
        self.setStyleSheet(f"""QMainWindow{{background:{C['base']}}QWidget{{background:{C['base']};color:{C['text']};font-family:'Segoe UI',sans-serif;}}
QMainWindow{}QWidget{{background:{C['base']}}QWidget{{background:{C['base']};color:{C['text']}}
QLineEdit,QTextEdit,QComboBox,QSpinBox{{background:{C['surface0']};color:{C['text']};border:1px solid {C['surface1']};border-radius:6px;padding:8px;font-size:12px;}}
QLineEdit:focus,QTextEdit:focus,QComboBox:focus{{border-color:{C['blue']}}
QComboBox::drop-down{{border:none;background:{C['surface1']};border-radius:0 6px 0;width:30px;}}
QComboBox QAbstractItemView{{background:{C['surface0']};color:{C['text']};selection-background-color:{C['surface1']};}}
QSpinBox::up-button,QSpinBox::down-button{{background:{C['surface1']};border:none;width:20px;}}
QCheckBox{{color:{C['text']};spacing:8px;}}
QCheckBox::indicator{{width:18px;height:18px;border-radius:4px;border:2px solid {C['surface2']};background:{C['surface0']};}}
QCheckBox::indicator:checked{{background:{C['blue']};border-color:{C['blue']};}}
QScrollBar:vertical{{background:{C['base']};width:8px;border-radius:4px;}}
QScrollBar::handle:vertical{{background:{C['surface1']};border-radius:4px;min-height:30px;}}QScrollBar::add-line:vertical,QScrollBar::sub-line:vertical{{height:0;}}
QSlider::groove:horizontal{{background:{C['surface1']};height:6px;border-radius:3px;}}
QSlider::handle:horizontal{{background:{C['blue']};width:16px;height:16px;margin:-5px 0;border-radius:8px;}}
QSlider::sub-page:horizontal{{background:{C['blue'];border-radius:3px;}}""")
    def _ui(self):
        cw=QWidget(); self.setCentralWidget(cw); ml=QVBoxLayout(cw); ml.setContentsMargins(0,0,0,0); ml.setSpacing(0)
        splitter=QSplitter(Qt.Vertical)
        top_w=QWidget(); hl=QHBoxLayout(top_w); hl.setContentsMargins(0,0,0,0); hl.setSpacing(0)
        sb=QFrame(); sb.setFixedWidth(220)
        sb.setStyleSheet(f"background:{C['mantle']};border-right:1px solid {C['surface0']};"); sl=QVBoxLayout(sb); sl.setContentsMargins(0,0,0,0)
        lo=QLabel("🤖 PresAgent v2.5"); lo.setFont(QFont("Segoe UI",13,QFont.Bold)); lo.setStyleSheet(f"color:{C['blue']};padding:16px 12px 8px;"); sl.addWidget(lo)
        self.si=Steps(self.steps); sl.addWidget(self.si,1)
        ctrl=QHBoxLayout(); save_btn=GhostBtn("💾 Save"); save_btn.clicked.connect(self.save_project); ctrl.addWidget(save_btn)
        load_btn=GhostBtn("📂 Load"); load_btn.clicked.connect(self.load_project); ctrl.addWidget(load_btn)
        ft=QLabel("Web Search Enabled • Executive Architect"); ft.setStyleSheet(f"color:{C['green']};font-size:9px;padding:4px;"); ft.setAlignment(Qt.AlignCenter); sl.addWidget(ft)
        hl.addWidget(sb); hl.addWidget(ca); cl=QVBoxLayout(ca); cl.setContentsMargins(0,0,0,0)
        self.stack=QStackedWidget(); cl.addWidget(self.stack)
        self.ov=Overlay(self.stack); cl.addWidget(self.ov)
        hl.addWidget(ca,1)
        splitter.addWidget(top_w)
        self.tool_log=ToolLogPanel(); self.tool_log.setVisible(False)
        splitter.addWidget(self.tool_log); splitter.setStretchFactor(0,4); splitter.setStretchFactor(1,1)
        ml.addWidget(splitter)
        self.pg_config=ConfigPage(); self.pg_brief=BriefPage(); self.pg_outline=OutlinePage(); self.pg_content=ContentPage(); self.pg_notes=NotesPage(); self.pg_export=ExportPage()
        for w in [self.pg_config,self.pg_brief,self.pg_outline,self.pg_content,self.pg_notes,self.pg_export]: self.stack.addWidget(w)
        self.statusBar().setStyleSheet(f"background:{C['mantle']};color:{C['subtext']};font-size:11px;border-top:1px solid {C['surface0']};")
        self.log_toggle=QPushButton("📡 Log")
        self.log_toggle.setCursor(QCursor(Qt.PointingHandCursor))
        self.log_toggle.setStyleSheet(f"background:{C['surface1']};color:{C['text']};border:none;border-radius:4px;padding:4px 12px;font-size:11px;")
        self.log_toggle.clicked.connect(self._toggle_log)
        self.statusBar().addPermanentWidget(self.log_toggle)
        self.statusBar().showMessage("Ready — Configure AI to begin")

    def _toggle_log(self):
        v=not self.tool_log.isVisible(); self.tool_log.setVisible(v)
        self.log_toggle.setText("Hide Log" if v else "📡 Log")

    def _connect(self):
        self.pg_config.ready.connect(self._go_brief)
        self.pg_brief.go.connect(self._gen_outline)
        self.pg_outline.approve.connect(self._gen_content)
        self.pg_outline.back.connect(lambda:self._go(1)); self.pg_outline.rg.clicked.connect(self._gen_outline)
        self.pg_content.go.connect(self._gen_notes)
        self.pg_content.back.connect(lambda:self._go(2)); self.pg_content.rg.clicked.connect(self._gen_content)
        self.pg_notes.go.connect(self._go_export)
        self.pg_notes.back.connect(lambda:self._go(3)); self.pg_notes.rg.clicked.connect(self._gen_notes)
        self.pg_export.back.connect(lambda:self._go(4))
        self.si.clicked.connect(lambda i: self._go(i) if i<=self.cur else None)
    def _go(self,i):
        self.cur=i; self.stack.setCurrentIndex(i); self.si.set_current(i)
        self.statusBar().showMessage(f"Step {i+1}/5 — {self.steps[i]}")
    def _go_brief(self):
        try: self.prov=self.pg_config.get_provider(); self._go(1)
        except ValueError as e: QMessageBox.warning(self,"Error",str(e))
    def _create_harness(self):
        self.tools=self.pg_config.get_tools(self.cfg)
        self.harness=SkillHarness(self.prov,self.tools,log_cb=lambda m: self.tool_log.append_log(m))
    def _gen_outline(self):
        self.cfg=self.pg_brief.get_config()
        if not self.cfg.topic: QMessageBox.warning(self,"Missing","Enter topic."); return
        self.ov.show_loading("Architecting outline..."); self.statusBar().showMessage("🤖 Architecting structure...")
        self._create_harness(); self.tool_log.append_log("🚀 Architect Skill (Web Search ENABLED" if self.cfg.enable_search else "")
        w=SkillWorker(self.prov,self.harness,OutlineSkill(),self.cfg)
        w.progress.connect(lambda m:self.statusBar().showMessage(f"🤖 {m}"))
        w.result.connect(self._on_outline); w.error.connect(self._err); w.done.connect(self.ov.hide_loading); w.start()
    def _on_outline(self,sl):
        self.slides=sl; self.pg_outline.set_slides(sl); self.si.done.add(1); self._go(2)
        self.tool_log.append_log(f"✅ {len(sl)} slides architected")
    def _gen_content(self):
        if not self.slides: return
        self.ov.show_loading("Writing visual content..."); self._create_harness(); self.tool_log.append_log("🚀 Visual Writer Skill")
        w=SkillWorker(self.prov,self.harness,ContentSkill(),self.cfg,self.slides)
        w.progress.connect(lambda m:self.statusBar().showMessage(f"🤖 {m}"))
        w.result.connect(self._on_content); w.error.connect(self._err); w.done.connect(self.ov.hide_loading); w.start()
    def _on_content(self,sl):
        self.slides=sl; self.pg_content.set_slides(sl); self.si.done.add(2); self._go(3)
        self.tool_log.append_log("✅ Visual content written")
    def _gen_notes(self):
        if not self.slides: return
        self.ov.show_loading("Directing stage notes..."); self._create_harness(); self.tool_log.append_log("🚀 Stage Director Skill")
        w=SkillWorker(self.prov,self.harness,NotesSkill(),self.cfg,self.slides)
        w.progress.connect(lambda m:self.statusBar().showMessage(f"🤖 {m}"))
        w.result.connect(self._on_notes); w.error.connect(self._err); w.done.connect(self._ov.hide_loading); w.start()
    def _on_notes(self,sl):
        self.slides=sl; self.pg_notes.set_slides(sl); self.si.done.add(3); self._go(4)
        self.tool_log.append_log("✅ Stage directions written")
    def _go_export(self):
        self.pg_export.update_stats(self.slides); self.pg_export.slides=self.slides; self.pg_export.cfg=self.cfg; self.si.done.add(4); self._go(5)
    def _err(self,msg):
        self.ov.hide_loading(); self.tool_log.append_log(f"❌ {str(msg)[:100]}")
        QMessageBox.critical(self,"Error",f"{msg}\n\nUse Demo Mode to test without API key.")
    def resizeEvent(self,e): super().resizeEvent(e); self.ov.setGeometry(self.stack.rect())
    def save_project(self):
        self._sync_slides_from_ui()
        p,_=QFileDialog.getSaveFileName(self,"Save Project","presentation.json","JSON (*.json)")
        if p:
            try:
                open(p,'w').write(json.dumps({"config":asdict(self.cfg),"slides":[asdict(s) for s in self.slides]},indent=2)); self.project_path=p; self.show_toast("Saved!", C['green'])
            except Exception as e: self.show_toast(f"Save Failed", C['red'])
    def load_project(self):
        p,_=QFileDialog.getOpenFileName(self,"Load Project","","JSON (*.json)")
        if p:
            try:
                data=json.load(open(p))
                self.cfg=PresConfig(**{k:v for k,v in data.get('config',{}).items() if k in PresConfig.__dataclass_fields__})
                self.slides=[SlideData(**{k:v for k,v in data.get('slides',[])} for k in SlideData.__dataclass_fields__})]
                self.project_path=p
                self.pg_brief.tp.setText(self.cfg.topic); self.pg_brief.au.setCurrentText(self.cfg.audience); self.pg_brief.ns.setValue(self.cfg.num_slides); self.pg_brief.tn.setCurrentText(self.cfg.tone)
                self.pg_outline.set_slides(self.slides); self.pg_content.set_slides(self.slides); self.pg_notes.set_slides(self.slides); self.pg_export.update_stats(self.slides)
                self.si.done.add(2); self._go(2); self.show_toast("Project Loaded!", C['blue'])
            except Exception as e: self.show_toast(f"Load Failed", C['red'])
    def _sync_slides_from_ui(self):
        if self.cur==3 and self.pg_content.tabs.count()==len(self.slides):
            for i,s in enumerate(self.slides):
                w=self.pg_content.tabs.widget(i)
                if w and w.layout().count()>=6: s.title=w.layout().itemAt(1).widget().text()
                if w and w.layout().count()>=8: s.body=w.layout().itemAt(3).widget().toPlainText()
                if w and w.layout().count()>=10: s.bullets=[l.strip() for l in w.layout().itemAt(5).widget().toPlainText().split('\n') if l.strip()]
                if w and w.layout().count()>=12: s.highlight=w.layout().itemAt(7).widget().text()
                if w and w.layout().count()>=14: s.visual_suggestion=w.layout().itemAt(9).widget().text()
    def _go_export(self):
        self.pg_export.update_stats(self.slides); self.pg_export.slides=self.slides; self.pg_export.cfg=self.cfg; self.si.done.add(4); self._go(5)
    def resizeEvent(self,e): super().resizeEvent(e); self.ov.setGeometry(self.stack.rect())


# ═══════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════════
def main():
    app=QApplication(sys.argv); app.setStyle("Fusion")
    pal=QPalette()
    pal.setColor(QPalette.Window,QColor(C['base'])); pal.setColor(QPalette.WindowText,QColor(C['text']))
    pal.setColor(QPalette.Base,QColor(C['surface0'])); pal.setColor(QPalette.Text,QColor(C['text']))
    pal.setColor(QPalette.Button,QColor(C['surface0'])); pal.setColor(QPalette.ButtonText,QColor(C['text']))
    pal.setColor(QPalette.Highlight,QColor(C['blue'])); pal.setColor(QPalette.HighlightedText,QColor(C['crust']))
    app.setPalette(pal)
    w=MainWindow(); w.show()
    sys.exit(app.exec())

if __name__=='__main__':
    main()