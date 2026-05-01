import io
import re
import json
import unicodedata
import html
from pathlib import Path
from collections import defaultdict, Counter
from zipfile import BadZipFile
import pandas as pd
from docx import Document
import streamlit as st
from datetime import datetime

# ============================================================
# Constants (Rulebook)
# ============================================================

CURRENT_YEAR = datetime.now().year
BY_CATEGORY_SHEET = "By Category"

REQUIRED_COLUMNS = [
    "Name",
    "Years of Experience",
    "Language",
    "Cloud",
    "Databases",
    "OS",
    "Framework/Libraries",
    "DevOps",
    "Container/Orchestration",
    "Machine Learning/AI",
    "Networking",
    "Version Control",
    "Tools",
    "Other",
    "Certifications",
    "Degree/Associates",
    "Degree/Bachelors",
    "Degree/Masters",
    "Degree/Phds"
]

SKILL_CATEGORY_COLS = [
    "Language",
    "Cloud",
    "Databases",
    "OS",
    "Framework/Libraries",
    "DevOps",
    "Container/Orchestration",
    "Machine Learning/AI",
    "Networking",
    "Version Control",
    "Tools",
    "Other"
]

SKILLS_HEADERS = {
    "skills",
    "technical skills",
    "technical skill",
    "skills/tools",
    "skills/tools/technologies",
    "skills & tools",
    "skills &amp; tools"
}


EDU_HEADERS = {"education", "educations"}
CERT_HEADERS = {"certification", "certifications"}

EXPERIENCE_HEADERS = {
    "professional experience",
    "experience",
    "work experience",
    "employment history",
    "work history",
    "experience (ts/sci full-scope cleared)"
}

NON_EXTRACTABLE_HEADERS = {
    
    "association/honors",
    "associations/honors",
    "associations",
    "honors",
    "summary",
    "clearance",
    "clearances",
    "projects",
    "project experience",
    "publications",
    "awards",
    "organizations",
    "activities",
    "references",
    "training",
    "trainings",
    "courses",
    "course",
    "membership",
    "memberships",
    "traininig", 
    "technologies",
    "technology",
    "tech",
    "skills/technologies",
    "skills/technology",
    "skills/tech",
    "training/courses",
    "independent projects",
    "home lab",
    "independent projects &amp; home lab"
    
    
}

_SPLIT = re.compile(r"[;,]")

# Paren-aware splitting and skill expansion (added to fix parser breaking on
# commas inside parentheses, e.g. "Linux (Ubuntu, Kali)" was being split into
# "Linux (Ubuntu" and "Kali)").

_QUALIFIER_RX = re.compile(
    r"\s*[-–—]\s*(daily use|primary|main|production|prod|development|dev|"
    r"preferred|backup|legacy|alternative|current|former|previous)\s*$",
    re.I,
)

_BARE_QUALIFIER_WORDS = {
    "daily", "primary", "main", "production", "prod", "development", "dev",
    "preferred", "backup", "legacy", "alternative", "current", "former",
    "previous", "self-taught", "ongoing", "in progress", "wip", "occasional",
    "occasionally", "rarely", "advanced", "intermediate", "beginner", "expert",
    "proficient", "familiar", "basic", "introductory", "familiarity", "segmentation", "authoring",
}

_VERSION_ONLY_RX = re.compile(r"^v?\d+(\.\d+)*\+?$")

def _paren_aware_split(s):
    """Split on commas and semicolons, but NOT when inside parentheses."""
    if not s:
        return []
    parts = []
    buf = []
    depth = 0
    for ch in s:
        if ch == "(":
            depth += 1
            buf.append(ch)
        elif ch == ")":
            depth = max(0, depth - 1)
            buf.append(ch)
        elif ch in ",;" and depth == 0:
            piece = "".join(buf).strip()
            if piece:
                parts.append(piece)
            buf = []
        else:
            buf.append(ch)
    piece = "".join(buf).strip()
    if piece:
        parts.append(piece)
    return parts


def _expand_skill_piece(piece):
    """Take one comma-split skill piece and expand it into one or more skills.

    - Extracts parenthetical content as additional skills.
    - Strips qualifier suffixes like "- daily use", "- primary".
    - Drops paren content that is just a bare qualifier or a version number.
    - Drops the parens from the main name.

    "Linux (Ubuntu, Kali - daily use)" -> ["Linux", "Ubuntu", "Kali"]
    "firewall configuration (Ubiquiti UDM Pro SE)" -> ["firewall configuration", "Ubiquiti UDM Pro SE"]
    "Bash (daily)" -> ["Bash"]
    "Python (3.10+)" -> ["Python"]
    "Kubernetes" -> ["Kubernetes"]
    """
    if not piece:
        return []

    out = []
    paren_items = []

    # Pull out content from each (...) group
    for m in re.finditer(r"\(([^)]*)\)", piece):
        inner = m.group(1)
        for inner_part in re.split(r"[,;]", inner):
            inner_part = inner_part.strip()
            inner_part = _QUALIFIER_RX.sub("", inner_part).strip()
            if not inner_part:
                continue
            # Drop bare qualifier words like "daily", "preferred"
            if inner_part.lower() in _BARE_QUALIFIER_WORDS:
                continue
            # Drop version-only content like "3.10+", "v2"
            if _VERSION_ONLY_RX.match(inner_part):
                continue
            paren_items.append(inner_part)

    # Main name: piece with parens removed
    main = re.sub(r"\s*\([^)]*\)", "", piece).strip()
    main = _QUALIFIER_RX.sub("", main).strip()
    if main:
        out.append(main)

    out.extend(paren_items)
    return out

_DASHES = str.maketrans({
    "‐": "-",
    "‑": "-",
    "‒": "-",
    "–": "-",
    "—": "-",
    "−": "-",
})

# ============================================================
# Rule 10: Compound splits (must run BEFORE normalization)
# ============================================================

COMPOUND_SPLITS = {
    "gitlab ci/cd": ["GitLab", "CI/CD"],
    "google earth/maps api": ["Google Earth", "Maps API"],
    "html/css": ["HTML", "CSS"],
    "oracle pl/sql": ["Oracle", "PL/SQL"],
    "scrum/kanban": ["Scrum", "Kanban"],
    "nifi gitlab": ["NiFI", "GitLab"],
    "mongodb spring jpa": ["MongoDB Spring", "Spring JPA"],
    "vmware java ee": ["VMWare", "Java EE"],
    "symantec. mcafee": ["Symantec", "McAfee"],
    "php visual basic": ["PHP", "Visual Basic"],
    "ida pro lod4j": ["IDA Pro", "Lod4j"],  
    "tomcat infinispan": ["Tomcat", "Infinispan"],
    "docker swarm docker compose": ["Docker Swarm", "Docker Compose"],
    "slick/scalaquery": ["SLICK", "ScalaQuery"],
    "xml/sqd" : ["XML","SQD"],
    "xml/xsd": ["XML", "XSD"],
    "analog/digital oscilloscope":["Analog Oscilloscope","Digital Oscilloscope"],
    "elasticsearch cvs": ["ElasticSearch", "CVS"],
    "windows server / 10 / 11": ["Windows Server 10", "Windows Server 11"],
    "802.1x / radius" : ["802.1X", "RADIUS"],
    "wireshark / tshark packet capture analysis": ["Wireshark","TShark packet capture analysis"],
    "microsoft sentinel / azure monitor": ["Microsoft Sentinel", "Azure Monitor"],
    "aws ec2": ["AWS", "EC2"],
    "ssp / poa&m / sprs authoring": ["SSP", "POA&M", "SPRS"],
    "embedded and parallel VxWorks": ["Embedded VxWorks",  "Parallel VxWorks"],
    "xlinx and altera": ["Xilinx", "Altera"]

}

# ============================================================
# Rule 11: Skill normalization (always applied)
# ============================================================

SKILL_NORMALIZATION = {
    
    "amazon aws": "AWS",
    "some aws": "AWS",
    "a ws": "AWS",
    "amazon": "AWS",
    "amazon ec2": "EC2",
    "ec2": "EC2",
    "arcgis tool": "ArcGIS",
    "arcgis tools": "ArcGIS",
    "bash scripting": "Bash",
    "extjs": "ExtJS",
    "ext.js": "ExtJS",
    "ida pro": "IDA Pro",
    "idapro": "IDA Pro",
    "jaws reader": "Jaws",
    "jupyter notebook": "Jupyter Notebooks",
    "jupyter": "Jupyter Notebooks",
    "red hat": "RedHat",
    "mips assembly": "Mips",
    "network protocol suites": "Network Protocol",
    "network protocols": "Network Protocol",
    "nodejs": "Node.js",
    "node.js": "Node.js",
    "objective c": "Objective‑C",
    "objective‑c": "Objective‑C",
    "onnx runtime": "ONNX",
    "python3": "Python",
    "r studio": "R",
    "and identity access management": "Identity Access Management",
    "groovy/grails tool suite": "Groovy/Grails",
    "groovy/grails tol suite": "Groovy/Grails",
    "groovy": "Groovy/Grails",
    "sql-lite": "SQLite",
    "vue.js": "Vue",
    "vue.js": "Vue",
    "visual studio code": "VSCode",
    "visual studio": "VSCode",
    "visual studio.": "VSCode",
    "vs code": "VSCode",
    "ms visual studio": "VSCode",
    "ajax": "AJAX",
    "randomforests": "Random Forests",
    "gaussian models.": "Gaussian Models",
    "nas and san storage arrays": "NAS/SAN Storage",
    "xlinx design suite": "Xilinx",
    "visio 2000": "Visio",
    "ms visio": "Visio",
    "microsoft office tools": "Microsoft Office",
    "xlinx and altera:": "FPGA Design",
    "all lsi and msi logic families": "LSI/MSI Logic",
    "spring boot deployments": "Spring Boot",
    "linux scripting": "LINUX",
    "linix": "LINUX",
    "linux servers": "LINUX",
    "cloudwatch": "CloudWatch",
    "res instances": "RES",
    "lambdas": "Lambda",
    "lambda": "Lambda",
    "ec2 instances": "EC2",
    "elasticsearch cvs": "Elasticsearch",
    "elasticsearch": "Elasticsearch",
    "google earthmaps api": "Google Earth",
    "apache nifi": "Apache NiFi",
    "cvs.": "CVS",
    "shell": "Shell",
    "nifi": "NiFi",
    "gitlab": "GitLab",
    "openvpn.": "OpenVPN",
    "vmware": "VMWare",
    "xp": "Windows XP",
    "vista": "Windows Vista",
    "ubuntu.": "Ubuntu",
    "mash vm user.": "Mash VM",
    "pki": "PKIs",
    "removeview.": "RemoveView",
    "clustering": "Clustering",
    "cluster computing": "Clustering",
    "bash scripting": "Bash Scripting",
    "hpc": "High Performance Computing",
    "rhel": "RHEL",
    "windows server.": "Windows Server",
    "windows servers":"Windows Server",
    "window": "Windows",
    "remoteview.": "RemoteView",
    "microsoft windows xp": "Windows XP",
    "tomcat": "Tomcat",
    "matploblib": "Matplotlib",
    "sgu hardware": "SGI",
    "microsoft": "Microsoft Office",
    "ssh protocols": "SSH",
    "dns server config": "DNS",
    "scrum": "SCRUM",
    "swager": "Swagger",
    "cisco works:": "Cisco",
    "cisco prime.": "Cisco",
    "cisco event scripting": "Cisco",
    "and mgx 8800 series atm switches": "MGX 8800 Series ATM Switches",
    "tcl.": "TCL",
    "juniper srx series.": "Juniper SRX series",
    "and kiribati.": "Kiribati",
    "and kiribiti.": "Kiribati",
    "agile methodology": "Agile",
    "apache http server": "Apache HTTP",
    "http.": "HTTP",
    "consul and vault.":"Consultaiton and Vault",
    "java spring cloud": "Java Spring",
    "junit4/5": "Junit",
    "junit": "Junit",
    "macosx": "MacOS X",
    "macos": "MacOS",
    "plsql": "PL/SQL",
    "spark.ml": "Spark",
    "sql developer": "SQL",
    "tensor analysis tool kit": "Tensor Analysis",
    "unix shell scripting": "UNIX",
    "ansible": "Ansible",
    "yml specs.":"YAML",
    ".bt templates":"bt templates",
    "x86 assembly": "x86",
    "x86dgb":"x86",
    "signal generator": "Signal Generator",
    "digital logic analyzer": "Digital Logic Analyzer",
    "sharepoint":"SharePoint",
    "java": "Java",
    "my sql": "MySQL",
    "rest apis": "REST API",
    "scikit-learn":"Scikit-learn",
    "tensorflow": "TensorFlow",
    "kali": "Kali Linux",
    "vlan segmentation": "VLAN",
    "typescript": "TypeScript",
    "ghidra": "GHIDRA",
    "nmap": "Nmap",
    "jira": "Jira",
    "scanboy": "ScanBoy",
    "cameo": "Cameo",
    "gimp": "Gimp",
    "centos": "CentOS",
    "centos": "CentOS",
    "bitbucket": "BitBucket",
    "ai":"Artificial Intelligence (AI)",
    "artificial intelligence": "Artificial Intelligence (AI)",
    "artificial intelligence (ai)": "Artificial Intelligence (AI)",
    "artificial intelligence (ai).": "Artificial Intelligence (AI)",
    "metasploit": "Metasploit",
    "mattermost": "Mattermost",
    "postgres": "PostGres",
    "vmware vcenter": "VMware vCenter",
    "strongswan ipsec": "StrongSwan IPsec",
    "powershell": "PowerShell",
    "elastic stack": "ElasticStack",
    "elasticstack": "ElasticStack",
    "ubiquity": "Ubiquiti"

    
}

 
# ============================================================
# Rule 12: Skill removal
# ============================================================

REMOVED_SKILLS = {""}

# ============================================================
# Rule 16: Certification normalization (CertificationFrequency ONLY)
# ============================================================

CERT_NORMALIZATION = {
    "aws solutions architect": "AWS Certified Solutions Architect - Associate",
    "aws solutions architect - associate (aws certified)": "AWS Certified Solutions Architect - Associate",
    "aws solutions architect associate": "AWS Certified Solutions Architect - Associate",
    "aws solutions architect – associate": "AWS Certified Solutions Architect - Associate",
    "aws solutions certified - associate": "AWS Certified Solutions Architect - Associate",
    "aws certified developer": "AWS Certified Developer - Associate",
    "aws solutions architect (in process)": "AWS Certified Solutions Architect - Associate",
    "self-study (udemy courses): aws cloud practitioner": "AWS Certified Cloud Practitioner",
    "ccna (cisco certified network associate)": "CCNA",
    "ccna certification": "CCNA",
    "itil v3 foundation": "ITIL v3.0",
    "iril v3 foundation (cert# 894862)": "ITIL v3.0",
    "security+": "CompTIA Security+",
    "security +": "CompTIA Security+",
    "comptia security+ (dod 8140 iat ii compliant)": "CompTIA Security+",
    "security+ (comptia)": "CompTIA Security+",
    "security+ ce": "CompTIA Security+",
    "comptia security+ ce": "CompTIA Security+",
    "comptia – security+": "CompTIA Security+",
    "comptia - security+": "CompTIA Security+",
    "comptia security+ ()":"CompTIA Security+",
    "security+ certification": "CompTIA Security+",
    "comptia networking +": "CompTIA Network+",
    "certified scrum master (csm)": "Certified Scrum Master",
    "certified scrum master": "Scrum Master",
    "certified scrum master - scrum alliance": "Scrum Master",
    "scrum alliance certified scrum master": "Scrum Master",
    "pmi agile certified practitioner (pmi-acp)":"PMI-ACP (Agile Certified Practitioner)",
    "pmp - project management institute (pmi)":"PMP (Project Management Professional)",
    "linux+ (comptia)": "CompTIA Linux+"
}

# ============================================================
# Rule 18: Degree normalization mapping (explicit only)
# ============================================================

def _degree_lookup_key(line: str) -> str:
    s = html.unescape(line or "")
    s = unicodedata.normalize("NFKC", s).translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    return s.lower()


# Your explicit entries (you said you manually lowercased long keys as needed)
DEGREE_NORMALIZATION = {
    "a.a., computer science": "A.S., Computer Science",
    "b.a. in computer science degree": "B.S., Computer Science",
    "b.a. in middle east studies / arabic": "B.S., Middle East Studies / Arabic",
    "b.a., mathematics/computer science": "B.S., Mathematics/Computer Science",
    "b.s. computer information systems, may 2011": "B.S., Computer Information Systems",
    "b.s. computer science": "B.S., Computer Science",
    "b.s. electrical engineering (1996)": "B.S., Electrical Engineering",
    "b.s. electrical engineering  (1996)": "B.S., Electrical Engineering",
    "b.s., science, electrical engineering": "B.S., Science, Electrical Engineering",
    "bachelors of science, electrical engineering": "B.S., Electrical Engineering",
    "b.s. in computer networks & cybersecurity": "B.S., Computer Networks and Cybersecurity",
    "b.s., cyber security": "B.S., Cybersecurity",
    "b.s., business and administration, 2018": "B.S., Business and Administration",
    "bachelor in business admin": "B.S., Business and Administration",
    "b.s., business administration": "B.S., Business and Administration",
    "b.s., computer science (minor: business administration), towson university, 12/2010": "B.S., Computer Science",
    "b.s., computer science - university of maryland baltimore county (1999)": "B.S., Computer Science",
    "b.s., computer science, college of computer": "B.S., Computer Science",
    "b.s., electrical and computer engineering (minor in german)": "B.S., Electrical and Computer Engineering",
    "b.s., information systems management, university of maryland baltimore county (1988)": "B.S., Information Systems Management",
    "ba english, bryn mawr college": "B.S., English",
    "bachelor of engineering in computer science": "B.S., Computer Science",
    "bachelor of science degree, electronic media engineering": "B.S., Electronic Media Engineering",
    "bachelor of science electrical engineering": "B.S., Electrical Engineering",
    "bachelor of science in computer and information science": "B.S., Computer and Information Science",
    "bachelor of science in history": "B.S., History",
    "bachelor of science, electrical engineering, bucknell university, lewisburg pa (1984)": "B.S., Electrical Engineering",
    "bs in cybersecurity and computer science, mount st. mary’s university, december 2021": "B.S., Cybersecurity and Computer Science",
    "bs, business administration": "B.S., Business and Administration",
    "bs, computer science": "B.S., Computer Science",
    "bs, computer science - cyber": "B.S., Computer Science - Cyber",
    "bs, computer science, university of baltimore, 1986": "B.S., Computer Science",
    "bs, cybersecurity": "B.S., Cybersecurity",
    "bs, economics": "B.S., Economics",
    "bs, information technology": "B.S., Information Technology",
    "2005: ba from the college of william and mary major in international studies and minor in computer science.": "B.S., International Studies and minor in Computer Science",
    "(in progress) bs in computer networks and security umuc": "B.S., Computer Networks and Security",
    "b.s., computer engineering technology": "B.S., Computer Engineering",
    "m.s.,data analytics": "M.S., Data Analytics",
    "master of business administration (mba)": "MBA",
    "master’s in business administration (mba)": "MBA",
    "mba, information technology": "MBA",
    "mba, university of baltimore, 1992": "MBA",
    "m.s.,computer science": "M.S., Computer Science",
    "m.s.,computer in science": "M.S., Computer Science",
    "master of science, computer science, johns hopkins university, baltimore md (1997)": "M.S., Computer Science",
    "master of science, electrical engineering": "M.S., Electrical Engineering",
    "master of science, electrical engineering, johns hopkins university, baltimore md (1989)": "M.S., Electrical Engineering",
    "master’s of science in geographic information systems": "M.S., Geographic Information Systems",
    "m.s., engineering science": "M.S., Engineering",
    "m.s., digital forensics and cyber investigations": "M.S., Digital Forensics and Cyber Investigation"
    
}    

# Canonicalize DEGREE_NORMALIZATION keys once (keeps ONE dict variable)
DEGREE_NORMALIZATION = {_degree_lookup_key(k): v for k, v in DEGREE_NORMALIZATION.items()}

# Deterministic cleanup patterns
_degree_year_rx = re.compile(r"\b(19\d{2}|20\d{2})\b")
_degree_month_rx = re.compile(
    r"\b(?:jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)\b",
    re.I
)
_degree_inst_rx = re.compile(r"\b(university|college|school|johns hopkins|anne arundel cc|institute|campus)\b", re.I)
_degree_minor_rx = re.compile(r"\bminor\b", re.I)

def _degree_generic_cleanup(line: str) -> str:
    """
    Deterministic cleanup (Rule 18.1):
    - remove parentheticals containing year/month/minor
    - remove trailing comma segments containing year/month/institution (incl CC/community college)
    - normalize prefixes: BA/BS/Bachelor->B.S., MS/Master->M.S.
    - fix double period: B.S.. -> B.S.
    - enforce comma format after abbreviations
    - remove punctuation-only segments like "." (prevents "B.S., ., X")
    """
    s = html.unescape(line or "")
    s = unicodedata.normalize("NFKC", s).translate(_DASHES).strip()
    if not s:
        return ""

    # remove parentheses containing year/month/minor
    def paren_repl(m):
        inner = m.group(1)
        if _degree_year_rx.search(inner) or _degree_month_rx.search(inner) or _degree_minor_rx.search(inner):
            return ""
        return "(" + inner + ")"
    s = re.sub(r"\(([^)]*)\)", paren_repl, s)

    # drop trailing comma segments that look like institution/date/location metadata
    parts = [p.strip() for p in s.split(",")]
    kept = []
    for i, p in enumerate(parts):
        if i == 0:
            kept.append(p)
            continue
        if _degree_year_rx.search(p) or _degree_month_rx.search(p) or _degree_inst_rx.search(p):
            break
        kept.append(p)
    s = ", ".join([k for k in kept if k]).strip(" ,")

    # Normalize prefixes to prevent duplicates
    s = re.sub(r"^AA\b", "A.S.", s, flags=re.I)
    s = re.sub(r"^A\.A\b", "A.S.", s, flags=re.I)
    s = re.sub(r"^AS\b", "A.S.", s, flags=re.I)
    s = re.sub(r"^A\.S\b", "A.S.", s, flags=re.I)
    s = re.sub(r"^Associate of\b", "A.S.", s, flags=re.I)
    s = re.sub(r"^Associates of\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Associate of,\b", "A.S.", s, flags=re.I)
    s = re.sub(r"^Associate in\b", "A.S.", s, flags=re.I)
    s = re.sub(r"^Associates in,\b", "A.S.", s, flags=re.I)
    s = re.sub(r"^Associates in,\b", "A.S.", s, flags=re.I)    
    
    s = re.sub(r"^BA\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^B\.A\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^BS\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^B\.S\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor of Arts\b", "B.S.", s, flags=re.I)  
    s = re.sub(r"^Bachelors of Science,\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelors of Science\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor of Science\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor of Science,\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelors in Science,\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelors in Science\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor in Science\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor in Science,\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor of Engineering\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor in Engineering\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelors in Engineering\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelors of\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelors of,\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor in\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelors in\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelors in,\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor in,\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor’s Degree\b", "B.S.", s, flags=re.I)
    s = re.sub(r"^Bachelor of\b", "B.S.", s, flags=re.I)
    

    s = re.sub(r"^MS\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^M\.S\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^M\.S\ in\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Master of Science\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Master’s Degree\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Master of Sciences,\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Masters of Science\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Masters of Science,\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Masters of Sciences,\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Master in Sciences,\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Masters in Science\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Masters in Science,\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Masters in Sciences,\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Master in\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Masters in\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Masters in,\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^Master in,\b", "M.S.", s, flags=re.I)
    s = re.sub(r"^M.S., in Science\b", "M.S.", s, flags=re.I)



    # Fix double period bug in abbreviations: B.S.. -> B.S.
    s = re.sub(r"\b(B\.S|M\.S|A\.S|A\.A|B\.A)\.\.", r"\1.", s, flags=re.I)
    s = re.sub(r"\.\.+", ".", s)  # safe within degree strings

    # Enforce comma formatting after abbreviations
    s = re.sub(r"^(B\.S\.|M\.S\.|A\.S\.|A\.A\.)\s*(?!,)", r"\1, ", s)
    s = re.sub(r"\s*,\s*", ", ", s)
    s = re.sub(r"\s+", " ", s).strip()

    # Remove punctuation-only comma segments ('.', '--', etc.)
    segs = [seg.strip() for seg in s.split(",")]
    segs = [seg for seg in segs if seg and seg not in {".", "..", "..."}]
    cleaned_segs = []
    for seg in segs:
        if re.fullmatch(r"[.\-–—_]+", seg):
            continue
        cleaned_segs.append(seg)
    s = ", ".join(cleaned_segs).strip()

    # Final comma normalization
    s = re.sub(r"\s*,\s*", ", ", s).strip(" ,")
    return s

# ============================================================
# Generalized “Label:” stripping inside SKILLS section (locked feature)
# ============================================================

_LABEL_ANYWHERE_RX = re.compile(r"(?i)(^|[\s,])([A-Za-z][A-Za-z0-9 /&\-\+]{0,60})\s*:\s*")

def _strip_skill_group_labels(line):
    s = (line or "").strip()
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s)
    s = s.translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    s = _LABEL_ANYWHERE_RX.sub(", ", s)
    s = re.sub(r"\s*,\s*", ", ", s)
    s = re.sub(r"(,\s*){2,}", ", ", s).strip(" ,")
    return s

def _skills_tokens_from_lines(skills_lines):
    """
    Convert raw SKILLS section lines into individual skill tokens.

    Rule enforcement:
    - Extract skills literally as written
    - Split ONLY on commas and semicolons
    - Do NOT split on '/' unless explicitly listed in COMPOUND_SPLITS
    - Parenthetical expansion is deterministic and explicit
    """

    tokens = []

    for line in skills_lines:
        cleaned = _strip_skill_group_labels(line)
        if not cleaned:
            continue

        # ✅ NO slash-based splitting here
        for piece in _paren_aware_split(cleaned):
            tokens.extend(_expand_skill_piece(piece))

    return tokens

# ============================================================
# Certification parsing + cleaning (as previously implemented)
# ============================================================

_months_rx = r"(?:jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)"
_re_year = re.compile(r"\b(19\d{2}|20\d{2})\b")
_re_mmddyyyy = re.compile(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b")
_re_month_year = re.compile(rf"\b{_months_rx}\b\s*\d{{0,2}}\s*,?\s*(19\d{{2}}|20\d{{2}})", re.I)
_re_paren = re.compile(r"\(([^)]*)\)")
_re_long_id = re.compile(r"\b[A-Z0-9]{8,}\b")
_re_code_like = re.compile(r"\b(?:COMP\d+|F\w{10,}|V\w{10,})\b", re.I)
_re_metadata_words = re.compile(r"\b(certification issued|Cert|Analyst# 10080|License|in progress|October|present|udemy courses|self-study|certification issues|issued|exp\.?|expires|Cert# 11073|expiration|taking test|attended Reinvent conference|202|in process)\b", re.I)
_re_expiry_tail = re.compile(r"\s*[-–—]\s*(exp\.?|expires?|expiration)\b.*$",re.I,)

def _canon_text(s: str) -> str:
    s = unicodedata.normalize("NFKC", (s or "")).translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _cert_lookup_key(s: str) -> str:
    return _canon_text(s).casefold()

def clean_cert_token(token: str) -> str:
    if token is None:
        return ""

    s = _canon_text(str(token))
    if not s:
        return ""

    # Strip bullets
    s = re.sub(r"^[•\-\u2022\t\s]+", "", s).strip()

    # ✅ Remove expiry tail BEFORE anything else
    s = _re_expiry_tail.sub("", s).strip()

    # Preserve original AFTER expiry removal
    original = s

    # Drop year-only junk
    if re.fullmatch(r"(19\d{2}|20\d{2})", s):
        return ""

    # Remove parentheticals ONLY if they contain metadata
    def paren_repl(m):
        inner = m.group(1)
        if (
            _re_year.search(inner)
            or _re_mmddyyyy.search(inner)
            or _re_month_year.search(inner)
            or _re_metadata_words.search(inner)
        ):
            return ""
        return "(" + inner + ")"

    s = _re_paren.sub(paren_repl, s)

    # Truncate at metadata words (issued, in progress, etc.)
    m = _re_metadata_words.search(s)
    if m:
        s = s[:m.start()].strip(" -;:,.\t")

    # Remove date/id noise
    s = _re_month_year.sub("", s)
    s = _re_mmddyyyy.sub("", s)
    s = _re_year.sub("", s)
    s = _re_code_like.sub("", s)
    s = _re_long_id.sub("", s)

    # Remove trailing generic words
    s = re.sub(
        r"\b(Certification|Certificate|course|certification|compliant|certificate|cert|Cert)\b\.?$",
        "",
        s,
        flags=re.I,
    ).strip()

    # Normalize whitespace/punctuation only
    s = re.sub(r"\s+", " ", s).strip(" ;,-")

    # If nothing meaningful remains, drop it
    if not s:
        return ""

    # ✅ Integrity guard (literal preservation only)
    meta_at_start = bool(
        _re_metadata_words.search(original)
        and _re_metadata_words.search(original).start() == 0
    )

    if " " in original and " " not in s and not meta_at_start:
        return original.strip()

    return s.strip()

def clean_certifications_from_lines(cert_lines):
    cleaned = []
    for line in cert_lines:
        line = _canon_text(line)
        if not line:
            continue
        parts = _paren_aware_split(line)
        for p in parts:
            c = clean_cert_token(p)
            if c:
                cleaned.append(c)
    return list(dict.fromkeys(cleaned))

# ============================================================
# Utilities
# ============================================================

def ensure_by_category_columns(df):
    for col in REQUIRED_COLUMNS:
        if col not in df.columns:
            df[col] = ""
    return df[REQUIRED_COLUMNS].copy()

def create_empty_by_category_df():
    return pd.DataFrame(columns=REQUIRED_COLUMNS)

def dedupe_preserve_order(items):
    seen = set()
    out = []
    for x in items:
        t = (x or "").strip()
        if not t:
            continue
        k = t.lower()
        if k not in seen:
            seen.add(k)
            out.append(t)
    return out

def _clean_header(s):
    # normalize header text for matching
    t = (s or "").strip()
    t = t.lstrip("•-*–—\t ").rstrip(":").strip()
    return t.lower()

def _is_all_caps_header(line):
    t = (line or "").strip()
    if not t or len(t) > 60:
        return False
    letters = re.sub(r"[^A-Za-z]", "", t)
    if len(letters) < 3:
        return False
    return (t.upper() == t) and any(ch.isalpha() for ch in t)


def _is_section_boundary(line):
    h = _clean_header(line)

    if (
        h in SKILLS_HEADERS
        or h in EDU_HEADERS
        or h in CERT_HEADERS
        or h in EXPERIENCE_HEADERS
        or h in NON_EXTRACTABLE_HEADERS
    ):
        return True
    
    return False


# ============================================================
# Skill pipeline (Rule 9–11)
# ============================================================

def apply_compound_splitting(skill):
    k = re.sub(r"\s+", " ", (skill or "").strip().lower()).strip()
    return COMPOUND_SPLITS.get(k, [(skill or "").strip()])

def normalize_skill(skill):
    raw = (skill or "").strip()
    k = re.sub(r"\s+", " ", raw.lower()).strip()
    if k in REMOVED_SKILLS:
        return None
    return SKILL_NORMALIZATION.get(k, raw)

def process_skills(skills_raw):
    skills_raw = dedupe_preserve_order(skills_raw)
    expanded = []
    for s in skills_raw:
        expanded.extend(apply_compound_splitting(s))
    normalized = []
    for s in expanded:
        ns = normalize_skill(s)
        if ns:
            normalized.append(ns)
    return dedupe_preserve_order(normalized)

def parse_and_normalize_skills_from_cell(cell):
    if cell is None or (isinstance(cell, float) and pd.isna(cell)):
        return []
    raw_tokens = []
    for piece in _paren_aware_split(str(cell)):
        raw_tokens.extend(_expand_skill_piece(piece))
    return process_skills(raw_tokens)

# ============================================================
# Persistent category map (canonicalize keys)
# ============================================================

def load_category_map(path: Path):
    """
    Loads skill->category mappings from JSON.
    Returns {} if missing or invalid JSON.
    """
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}

def save_category_map(path: Path, mapping: dict):
    """
    Merge-safe persistence:
    - Keeps any existing mappings on disk that are missing from `mapping`
    - Applies updates from `mapping` (explicit new/changed choices win)
    - Prevents manual JSON edits from being lost across runs
    """
    path.parent.mkdir(parents=True, exist_ok=True)

    # Load existing mappings (if any) WITHOUT losing them
    existing = {}
    if path.exists():
        try:
            existing = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            existing = {}

    # Merge: existing keys remain unless overridden by new mapping
    merged = dict(existing)
    merged.update(mapping)

    path.write_text(
        json.dumps(merged, indent=2, ensure_ascii=False),
        encoding="utf-8"
    )

# ============================================================
# DOCX parsing (safe + preserves internal line breaks)
# ============================================================

def _docx_lines(file_like):
    try:
        try:
            file_like.seek(0)
        except Exception:
            pass
        doc = Document(file_like)
        out = []
        for p in doc.paragraphs:
            raw = (p.text or "").strip()
            if not raw:
                continue
            out.extend([x.strip() for x in raw.split("\n") if x.strip()])
        return out
    except (BadZipFile, Exception):
        return []

def peek_name_from_docx(file_like):
    for line in _docx_lines(file_like):
        if not _is_section_boundary(line):
            return line.strip()
    return ""

def _section_lines(lines, header_set, section_name=None):
    start = None
    for i, line in enumerate(lines):
        if _clean_header(line) in header_set:
            start = i + 1
            break
    if start is None:
        return []

    out = []
    for line in lines[start:]:
        # stop only on REAL section boundaries
        if _is_section_boundary(line):
            break

        out.append(line)

    return out

def parse_resume_sections(file_like, name_override=None):
    file_like.seek(0)
    lines = _docx_lines(file_like)

    name = (name_override or "").strip() or peek_name_from_docx(file_like)

    skills_lines = _section_lines(lines, SKILLS_HEADERS)
    edu_lines = _section_lines(lines, EDU_HEADERS)
    cert_lines = _section_lines(lines, CERT_HEADERS)
    exp_lines = _section_lines(lines, EXPERIENCE_HEADERS)

    skills_raw = dedupe_preserve_order(_skills_tokens_from_lines(skills_lines))

    # Rule 15: certifications extracted only from Certifications section, split on commas/semicolons, literal.
    # cert_text = " ".join(cert_lines).strip()
    # certs_raw = dedupe_preserve_order([c.strip() for c in _SPLIT.split(cert_text) if c.strip()])  
    certs_raw = clean_certifications_from_lines(cert_lines)


    return {
        "name": name,
        "skills_raw": skills_raw,
        "certs_raw": certs_raw,
        "education_lines": edu_lines,
        "education_text": " ".join(edu_lines).strip(),
        "experience_lines": exp_lines,
        "experience_text": " ".join(exp_lines).strip(),
    }


# Experience -> Oldest Job Year (Rule 8)

_MONTHS_NAME_RX = r"(?:jan|january|feb|february|mar|march|apr|april|may|jun|june|jul|july|aug|august|sep|sept|september|oct|october|nov|november|dec|december)"

def extract_oldest_experience_year(experience_lines):
    """
    Rule 8:
    Years of Experience = 2026 − year of oldest job experience.

    Supported explicit formats (start year only):
      - July 2019 – March 2022
      - Jul 2019 – Mar 2022
      - July, 2019 – March, 2022
      - July/2019 – March/2022
      - 5/1984-10/2021
      - MM/YYYY
      - YYYY

    No inference. Deterministic.
    """

    if not experience_lines:
        return None

    years = []

    for raw in experience_lines:
        s = html.unescape(raw or "")
        s = unicodedata.normalize("NFKC", s).translate(_DASHES)

        # ------------------------------------------------------------
        # Month YYYY – Month YYYY
        # July 2019 – March 2022
        # July, 2019 – March, 2022
        # ------------------------------------------------------------
        for y in re.findall(
            rf"\b{_MONTHS_NAME_RX}\b\s*,?\s+(19\d{{2}}|20\d{{2}})\s*[-–]\s*\b{_MONTHS_NAME_RX}\b\s*,?\s+\d{{4}}\b",
            s,
            flags=re.I
        ):
            years.append(int(y))

        # ------------------------------------------------------------
        # Abbreviated Month YYYY – Abbreviated Month YYYY
        # Jul 2019 – Mar 2022
        # ------------------------------------------------------------
        for y in re.findall(
            r"\b(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\b\s*,?\s+(19\d{2}|20\d{2})\s*[-–]\s*(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\b\s*,?\s+\d{4}\b",
            s,
            flags=re.I
        ):
            years.append(int(y))

        # ------------------------------------------------------------
        # Month/YYYY – Month/YYYY
        # July/2019 – March/2022
        # ------------------------------------------------------------
        for y in re.findall(
            rf"\b{_MONTHS_NAME_RX}/(19\d{{2}}|20\d{{2}})\s*[-–]\s*\b{_MONTHS_NAME_RX}/\d{{4}}\b",
            s,
            flags=re.I
        ):
            years.append(int(y))

        # ------------------------------------------------------------
        # MM/YYYY-MM/YYYY
        # 5/1984-10/2021
        # ------------------------------------------------------------
        for y in re.findall(
            r"\b(?:0?[1-9]|1[0-2])/(19\d{2}|20\d{2})\s*-\s*(?:0?[1-9]|1[0-2])/\d{4}\b",
            s
        ):
            years.append(int(y))

        # ------------------------------------------------------------
        # Standalone MM/YYYY
        # 01/2014
        # ------------------------------------------------------------
        for y in re.findall(
            r"\b(?:0?[1-9]|1[0-2])/(19\d{2}|20\d{2})\b",
            s
        ):
            years.append(int(y))

        # ------------------------------------------------------------
        # Standalone Month YYYY
        # July 2019
        # ------------------------------------------------------------
        for y in re.findall(
            rf"\b{_MONTHS_NAME_RX}\b\s+(19\d{{2}}|20\d{{2}})\b",
            s,
            flags=re.I
        ):
            years.append(int(y))

        # ------------------------------------------------------------
        # Standalone YYYY
        # 2014
        # ------------------------------------------------------------
        for y in re.findall(
            r"\b(19\d{2}|20\d{2})\b",
            s
        ):
            years.append(int(y))

    if not years:
        return None

    return min(years)

# ============================================================
# Mapping from existing By Category
# ============================================================

def build_category_map_from_by_category(by_cat):
    by_cat = ensure_by_category_columns(by_cat)
    mapping = {}
    conflicts = defaultdict(set)

    for _, row in by_cat.iterrows():
        for col in SKILL_CATEGORY_COLS:
            for s in parse_and_normalize_skills_from_cell(row.get(col, "")):
                if s in mapping and mapping[s] != col:
                    conflicts[s].update({mapping[s], col})
                else:
                    mapping[s] = col

    conflicts = {k: v for k, v in conflicts.items() if len(v) > 1}
    return mapping, conflicts

def resolve_conflicts_with_user(conflicts, category_map): 
    st.error("Conflicts found: same skill appears under multiple columns.")
    for skill in sorted(conflicts.keys(), key=lambda x: x.lower()):
        options = sorted(list(conflicts[skill]))
        choice = st.selectbox(
            f"Resolve category for skill: {skill}",
            options=options,
            key=f"conflict_{skill}",
        )
        category_map[skill] = choice
    st.warning("Resolve all conflicts, then rerun.")
    st.stop()
    return category_map

def categorize_skills_with_user(skills, category_map, resume_label=""):
    skills = process_skills(skills)
    unknown = [s for s in skills if s not in category_map]
    if unknown:
        st.warning(f"Uncategorized skills found in {resume_label}. Assign each to one category.")
        placeholder = "-- Select category --"
        options = [placeholder] + SKILL_CATEGORY_COLS

        incomplete = False
        for s in unknown:
            choice = st.selectbox(
                f"Select category for skill: {s}",
                options=options,
                index=0,
                key=f"cat_{resume_label}_{s}",
            )
            if choice == placeholder:
                incomplete = True
            else:
                category_map[s] = choice

        if incomplete:
            st.warning("Please select a category for every uncategorized skill.")
            st.stop()

    categorized = {c: [] for c in SKILL_CATEGORY_COLS}
    for s in skills:
        categorized[category_map[s]].append(s)

    for c in categorized:
        categorized[c] = sorted(categorized[c], key=lambda x: x.lower())

    return categorized, category_map

# ============================================================
# Degree extraction + cleanup mapping (Rule 18.1)
# ============================================================

def extract_degree_lines(edu_lines):
    #rx = re.compile(r"\b(Ph\.?D|PhD|Doctor|M\.?S|MBA|Master|B\.?S|B\.?A|BA|Bachelor|A\.?S|A\.?A|Associate)\b", re.I)
    rx = re.compile(
    r"(?:\b(Ph\.?D|PhD|Doctor|M\.?S|MBA|Master|Masters|Master's|B\.?S|B\.?A|BA|Bachelor|Bachelors|Bachelor's|Associate)\b|\bA\.\s*[SA]\.(?=\s|,|$))",
    re.I
)   
    out = []
    for line in edu_lines:
        if rx.search(line):
            out.append(line.strip())
    return dedupe_preserve_order(out)

def classify_degree(deg):
    d = (deg or "")
    low = d.lower()

    if re.search(r"\bph\.?d\b|\bdoctor\b", low):
        return "Degree/Phds"

    if re.search(r"\bm\.?s\.?\b|\bms\b|\bmba\b|\bmasters\b|\bmaster\b", low):
        return "Degree/Masters"
    
    if re.search(r"\bb\.?s\.?\b|\bb\.?a\.?\b|\bbachelor\b|\bbachelors\b", low):
        return "Degree/Bachelors"
    
    if re.search(r"\ba\.?s\.?\b|\ba\.?a\.?\b|\bassociate\b", low):
        return "Degree/Associates"

    # default bachelors if not above
    return "Degree/Bachelors"

def apply_degrees(edu_lines, edu_text):
    degrees_by_col = {
        "Degree/Associates": [],
        "Degree/Bachelors": [],
        "Degree/Masters": [],
        "Degree/Phds": [],
    }

    for deg in extract_degree_lines(edu_lines):
        # Step 1: canonical lookup key
        key = _degree_lookup_key(deg)

        # Step 2: explicit normalization mapping (if exists)
        norm = DEGREE_NORMALIZATION.get(key)

        # Step 3: deterministic cleanup if not explicitly mapped
        if norm is None:
            norm = _degree_generic_cleanup(deg)

        # Step 4: classify cleaned degree
        col = classify_degree(norm)
        degrees_by_col[col].append(norm)

    return degrees_by_col

# Compute Years of Experience

def compute_years_experience(oldest_job_year):
    if oldest_job_year is None:
        return None
    return CURRENT_YEAR - oldest_job_year

# ============================================================
# Upsert row (Rule 5)
# ============================================================

def upsert_candidate_row(by_cat, name, skills_by_category, certs_raw, degrees_by_col, oldest_job_year, action):
    by_cat = ensure_by_category_columns(by_cat)
    name_norm = (name or "").strip()
    exists_mask = by_cat["Name"].astype(str).str.lower() == name_norm.lower()

    if action == "replace" and exists_mask.any():
        by_cat = by_cat.loc[~exists_mask].copy()

    row = {c: "" for c in REQUIRED_COLUMNS}
    row["Name"] = name_norm
    row["Years of Experience"] = compute_years_experience(oldest_job_year)

    for col in SKILL_CATEGORY_COLS:
        row[col] = ", ".join(skills_by_category.get(col, []))

    row["Certifications"] = "; ".join(dedupe_preserve_order(certs_raw))

    for col in ["Degree/Associates", "Degree/Bachelors", "Degree/Masters", "Degree/Phds"]:
        row[col] = "; ".join(dedupe_preserve_order(degrees_by_col.get(col, [])))

    return pd.concat([by_cat, pd.DataFrame([row])], ignore_index=True)

# ============================================================
# Frequency rebuilds
# ============================================================

def _skill_freq_key(s):
    s = unicodedata.normalize("NFKC", (s or "").strip()).translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def rebuild_skill_frequency(by_cat):
    by_cat = ensure_by_category_columns(by_cat)
    counts = Counter()
    label_votes = defaultdict(Counter)

    for _, r in by_cat.iterrows():
        per_candidate = set()
        for col in SKILL_CATEGORY_COLS:
            for s in parse_and_normalize_skills_from_cell(r.get(col, "")):
                k = _skill_freq_key(s)
                if k:
                    per_candidate.add(k)
                    label_votes[k][s] += 1
        for k in per_candidate:
            counts[k] += 1

    key_to_label = {}
    for k, votes in label_votes.items():
        key_to_label[k] = sorted(votes.items(), key=lambda kv: (-kv[1], kv[0].casefold(), kv[0]))[0][0]

    rows = []
    for k in sorted(counts.keys(), key=lambda x: key_to_label.get(x, x).casefold()):
        rows.append({"Skill": key_to_label.get(k, k), "Candidate Count": int(counts[k])})

    return pd.DataFrame(rows, columns=["Skill", "Candidate Count"])

def _cert_key(s):
    s = unicodedata.normalize("NFKC", (s or "").strip()).translate(_DASHES)
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def _normalize_cert_for_frequency(cert):
    raw = (cert or "").strip()
    if not raw:
        return ""

    key = _cert_key(raw)

    # ✅ Rule 16: normalization ONLY for frequency sheet
    return CERT_NORMALIZATION.get(
        key,
        CERT_NORMALIZATION.get(raw.lower().strip(), raw)
    )

def rebuild_cert_frequency(by_cat):
    by_cat = ensure_by_category_columns(by_cat)
    counts = Counter()
    label_votes = defaultdict(Counter)

    for _, r in by_cat.iterrows():
        cell = r.get("Certifications", "")
        if not cell or not str(cell).strip():
            continue

        certs = _paren_aware_split(str(cell))
        per_candidate = set()

        for c in certs:
            norm = _normalize_cert_for_frequency(c)
            if not norm:
                continue
            k = _cert_key(norm)
            per_candidate.add(k)
            label_votes[k][norm] += 1

        for k in per_candidate:
            counts[k] += 1

    key_to_label = {
        k: sorted(v.items(), key=lambda kv: (-kv[1], kv[0].casefold()))[0][0]
        for k, v in label_votes.items()
    }

    rows = [
        {"Certification": key_to_label[k], "Candidate Count": counts[k]}
        for k in sorted(counts, key=lambda x: key_to_label[x].casefold())
    ]

    return pd.DataFrame(rows, columns=["Certification", "Candidate Count"])


def rebuild_degree_frequency(by_cat, col):
    by_cat = ensure_by_category_columns(by_cat)
    counts = Counter()

    for _, r in by_cat.iterrows():
        cell = r.get(col, "")
        if cell is None or (isinstance(cell, float) and pd.isna(cell)) or not str(cell).strip():
            continue

        degrees = [d.strip() for d in str(cell).split(";") if d.strip()]
        for d in set(degrees):
            counts[d] += 1

    return pd.DataFrame(sorted(counts.items(), key=lambda x: x[0].lower()),
                        columns=["Degree", "Candidate Count"])

# ============================================================
# Excel I/O
# ============================================================

def write_excel_output(by_cat, skill_freq, cert_freq, deg_assoc, deg_bach, deg_mast, deg_phd):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        ensure_by_category_columns(by_cat).to_excel(writer, sheet_name="By Category", index=False)
        skill_freq.to_excel(writer, sheet_name="SkillFrequency", index=False)
        cert_freq.to_excel(writer, sheet_name="CertificationFrequency", index=False)
        deg_assoc.to_excel(writer, sheet_name="DegreeFrequency_Associates", index=False)
        deg_bach.to_excel(writer, sheet_name="DegreeFrequency_Bachelors", index=False)
        deg_mast.to_excel(writer, sheet_name="DegreeFrequency_Masters", index=False)
        deg_phd.to_excel(writer, sheet_name="DegreeFrequency_Phds", index=False)
    return out.getvalue()

def load_excel_by_category(excel_file):
    xls = pd.ExcelFile(excel_file, engine="openpyxl")
    df = xls.parse(BY_CATEGORY_SHEET)
    return ensure_by_category_columns(df)