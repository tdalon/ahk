---
permalink: /Atlasy
title: "Atlasy"
excerpt: "Atlasy is a tool that extends Atlassian-based tools (like Jira, Confluence, R4J) capability."
redirect_from:
  - /atlasy
---


- [About](#about)
- [Main Tool Page](#main-tool-page)
- [Commands/ Keywords](#commands-keywords)
  - [Main Keywords](#main-keywords)
  - [Jira](#jira)
    - [Jira Quick Search](#jira-quick-search)
  - [Confluence](#confluence)
    - [Confluence Quick Open](#confluence-quick-open)
  - [R4J](#r4j)
  - [Xray](#xray)
    - [Doc](#doc)
  - [BigPicture](#bigpicture)
  - [BitBucket](#bitbucket)
- [Changelog](#changelog)


## About

Atlasy is a tool that enables your working within the Atlassian-based tools like Jira, Confluence, R4J etc.
It adds some UX related features and also extended capability.
You can run its features from an integrated Launcher with natural keywords and commands.

Short Feature highlights include:
 * bulk linking between Jira issues
 * quick search (Jira, Confluence)
 * quick navigation (e.g. R4J, Xray)
 * quick set an epic
 * ...


## [Main Tool Page](https://tdalon.blogspot.com/2023/09/atlasy-new-powertool.html)


## Commands/ Keywords

This is implemented in the main associated library [Lib/Atlasy.ahk](../Lib/Atlasy.ahk).

In the source code you can find the full up to date syntax for keywords and command. (I hope the code is clear enough.)

### Main Keywords

Keyword  |  Action
--|--
c  |  [Confluence](#confluence)
j   | [Jira](#jira)  
r   |  [R4J](#r4j)  
x | [Xray](#xray)
bb | [BitBucket](#bitbucket)
sw | switch between server and cloud (bidirectional)

<a name="jira"></a>
### Jira

Launcher Primary Keyword: 'j'

Secondary Keyword  |  Action
--|--
| If R4J window, open current Issue in Jira<br>
Open Selected or current Issues or Jira Root Url (if none)
Key | Open Project or Issue
-dp <projectKey>* |  Edit / Set Default Project  
-p <projectKey>* |  Navigate to Project
-pl  |  Edit Project List
-s | Open New Search window
-s or -i <Jql> | Open new search (Issues) prefilled with Jql. Supports Quick Search shortcuts.
-b <searchString> | Open Boards
-f <searchString> | Open Filters
-l or l | Add Link
-vl or vl | View Linked Issues
b | bulk edit selected or current issues
n | open selected or current issues in Navigator / Filter view
-c or c <projectKey>* | Open Full window Create Issue Screen. Project Key will be prefilled if passed as optional argument
h or -h  <keyword>* |  Open Help for command (if implemented)
?  | Open list of commands (this page)  

<a name="jira-quick-search"></a>
#### Jira Quick Search

Keyword  |  Jql
--|--
s~ | summary~
d~ | description~
-a | assigned to me (assignee = currentUser())
-c | created by me (creator = currentUser())
-u | unresolved (resolution = Unresolved)
-ua | unassigned (assignee is EMPTY)
-r | reported by me (reporter = currentUser())
-w | watched by me (watcher = currentUser())
#(label) | labels = label

<a name="confluence"></a>
### Confluence

Launcher Keyword: 'c'

Keywords  |  Action
--|--
| Open Confluence Root Url
$space | Open Confluence Space at Home
$space $query | Search in Space for query
o or -o $query | [Quick Open](#Quick Open)
o  | Re_order, View in Hierarchy
a  | view attachments
h  | view page history
i  | view page info

<a name="confluence-quick-open"></a>
#### Confluence Quick Open

s or -s followed by <spacekey> to restrict search in a Space
query: use # for search by label. You can combine multiple labels e.g. #label1 #label2 ; it will then search by AND i.e. intersection.
query can contain also keywords

Quick Open will open the first page found by the Search.

<a name="r4j"></a>
### R4J

Main keyword 'r'

Second Keyword or Command  |  Action
--|--
|  from Jira issue detailed view-> Open issue in R4J Tree
 ProjectKey or IssueKey $view| Open Issue or Project R4J view
 $view: d (document, tree) or c (coverage) or t (traceability)
-cp or cp | Copy Path Jql
-cc or cc | Copy Children Jql
-p or p | Transform and paste server Jql to Cloud Jql
-n or n | open in Issue navigator r4j path
-cv or -tv | Coverage or Traceability View commands
&nbsp;&nbsp;&nbsp;  c (default) | Copy
&nbsp;&nbsp;&nbsp;  i | Import
&nbsp;&nbsp;&nbsp;  e | Export


### Xray

Main keyword 'x'

2nd Keyword  |  Action
--|--
 | Open Xray Getting Started page
r   |  Test repository
e  |   Test execution
p   |  Test plans
m   | Test Plans Metrics
trace | Traceability Report
t | Tests List Report
ts | Test Sets Report
tp | Test Plan Report
te | Test Executions Report
tr | Test Run Report
[doc](#doc) | Open Xray documentation
gs  | Open Xray Getting Started page

#### Doc

3nd keyword | Page  
--|--
| Open main documentation page
rn | Release Notes
gs | Getting Started


<a name="bigpicture"></a>
### BigPicture

Main keyword is 'bp'

Opens BigPicture.

<a name="bitbucket"></a>
### BitBucket

Main keyword: 'bb'

<hr>

<a name="changelog"></a>
## Changelog

See [Atlasy Changelog](Atlasy-Changelog.md)
