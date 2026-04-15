---
title: 'FAIRspreadsheets: An Open-Source Application to Facilitate the Curation of Spreadsheet-based Research Data'
tags:
  - Python
  - research data curation
  - spreadsheets
  - FAIR data
  - data quality
authors:
  - name: Ellery Galvin
    orcid: 0009-0006-9435-4046 
    affiliation: 1
  - name: Alexandra Provo
    orcid: 0000-0001-8108-2025
    affiliation: 2
  - name: Aditya Ranganath
    orcid: 0000-0002-4721-2313
    affiliation: 1
  - name: Andrew Johnson
    orcid: 0000-0002-7952-6536
    affiliation: 1
  - name: Matthew Murray
    orcid: 0000-0001-5799-8471
    affiliation: 1
affiliations:
  - name: University of Colorado Boulder
    index: 1
    ror: 02ttsq026
  - name: New York University
    index: 2
    ror: 0190ak572
date: 13 April 2026
bibliography: paper.bib
---

# Summary

*FAIRspreadsheets* is a general-purpose, open-source application that helps data curators and researchers 
assess the quality and publication-readiness of spreadsheet-based research data. It implements a suite of programmatic tests that 
evaluate spreadsheets against various design and quality benchmarks 
suggested by the normative Information Science literature on 
spreadsheet data best practices, and generates a report that identifies areas in 
which a dataset departs from these quality benchmarks. Researchers, 
curators, and other research staff can use this report as a curatorial aid to evaluate 
a spreadsheet dataset’s alignment with widely accepted quality standards, and identify opportunities 
for improvement. More generally, the adoption of
*FAIRspreadsheets* has the potential to increase the efficiency and transparency of data curation workflows, and support the health of the FAIR data and reproducible research ecosystems. 

# Statement of need

Data publication is an increasingly widespread practice driven by 
evolving discipline-specific standards for open science and 
reproducibility [@christensen_transparent_2019; @powers_open_2019], 
journal policies [@bloom_data_2014; @piwowar_review_2008], and 
funder mandates [@martone_changing_2022; @nelson_ensuring_2022]. 
The FAIR data principles articulate broad guidelines for published 
research data, and a robust data publication ecosystem has evolved 
to ensure that data is Findable, Accessible, Interoperable, and 
Reusable [@wilkinson_fair_2016]. Research data curation is the 
process of reviewing and revising research data and related 
materials (such as metadata and documentation) to ensure that they 
are FAIR-compliant [@johnston_data_2017].

Despite the efforts of research institutions to build robust and 
well-staffed curation programs, it is difficult to scale labor-intensive curation 
activities to match the growth in research data publications. One way to address the challenge of scaling curation services
is to develop research data curation software that increases 
efficiency by automating aspects of the curatorial 
review process. To the extent that spreadsheets are among the most 
versatile and widely used tools to organize and store research data, software 
that facilitates the curation of spreadsheet-based research data is a particularly important need 
for researchers and research support professionals, especially as such datasets grow in 
size and complexity.

More specifically, information scientists, curators, and data professionals 
have articulated various quality benchmarks and best practices that enhance the 
reusability of spreadsheet-based research data 
[@briney_research_2023; @broman_data_2018; @houston_releasing_2014; 
@janee_microsoft_2019; @white_nine_2013]. When curating such datasets for publication, 
researchers and curators use these guidelines to inform their review and 
suggest recommendations for improvement. The efficiency, consistency, and thoroughness of the review process 
could be increased if researchers and curators have access to software that systematically identifies 
where and how spreadsheets diverge from these benchmarks, and flags areas for more focused
curatorial scrutiny. 

# State of the field

Despite the utility of such a software application for improving the quality of published research data, it has not been the 
focus of previous software development efforts. Existing spreadsheet feature 
detection tools tend to focus narrowly on specific spreadsheet properties, 
such as layouts [@vitagliano_detecting_2021], versioning 
[@xu_spreadcluster_2017], calculation models and formulas 
[@de_vos_methodology_2015; @jansen_detecting_2018], and tables 
[@jansen_use_2018; @dong_tablesense_2019].

@barik_fuse_2015 developed more general software to summarize certain spreadsheet 
properties, but these properties are not explicitly tied to spreadsheet best practices and quality benchmarks, which limits its utility as a curatorial tool. 
The Excel Archival Tool [@mcgrory_excel_2015] 
and the Qualitative Data Repository (QDR) Curation Handbook [@demgenski_introducing_2021]
are both examples of tools that support curation, with a focus on automating 
the conversion of spreadsheet files into open formats suitable for
long-term archiving and preservation. However, neither of these tools focuses
explicitly on intrinsic quality issues that are relevant before
a spreadsheet is converted to a file format appropriate for long-term preservation. Indeed, *FAIRspreadsheets*
is unique in meeting the need for a general-purpose feature detection tool focused 
on compatibility with a broad set of widely-accepted quality benchmarks. 

# Software design

*FAIRspreadsheets* is written in Python and uses only standard dependencies for tabular data analysis: `pandas` [@reback2020pandas], `numpy` [@harris2020array], and `openpyxl` [@openpyxl_software]. It implements a series of tests to ascertain whether a given dataset departs from specific best practices for spreadsheet data.  
  
Examples of the tests include:  
  
- Are there multiple tables on the same sheet?  
- Are there dates in formats other than YYYY-MM-DD?  
- Are there cells containing only question marks?  
- Are there special characters or leading/trailing spaces in cells?  
- Are there non-unique column headers?  
  
The architecture is modular, with each test implemented as a class that inherits most of the structure and functionality from abstract classes. There is a standardized interface through which tests can depend on one another that guarantees a network of tests is evaluated in the correct order and that all dependencies are included in the stack without explicit specification. Error handling ensures that runtime errors are caught without disrupting non-dependent tests, which is essential for large datasets.  

Adding a new test typically does not require detailed knowledge of the whole code base; it only requires specifying the test type, some metadata, and a boolean method returning the test result given its inputs and dependencies. For the user, options for each test can be optionally specified in a JSON file. This choice will keep the user’s code simple and allow them to programmatically generate different options from their preferred language. Eventually, a front end will generate these files from the user’s input so they can track which options led to which results.  
  
A separate layer of the architecture harvests the test results requested and generates reports; new tests created by the programmer automatically inherit this functionality. Users can pick any subset of the tests they wish to implement and designate either JSON or CSV as the format for the output report. These outputs include one object/row per failed test along with identifiers for the location of the issue, a description of the issue, and the content containing the issue. The output is both human readable and useful for pushing to a database, which may be useful for researchers aiming to characterize large collections of datasets. The software does not change the actual spreadsheet; rather, it guides the user in their own curatorial review, and is therefore  aligned with the paradigm of "human in the loop curation" [@demgenski_introducing_2021].

# Research impact

While designed as a practical curatorial tool for research data, we also expect 
*FAIRspreadsheets* to be a useful research tool, particularly in fields 
such as Library and Information Science. For example, the application could 
facilitate empirical studies exploring topics such as spreadsheet design,
data quality, data curation, and research data management.

# AI usage disclosure
Line completions, inline editing, and agents (Claude 4.5) provided by Cursor were used to aid in the development of this software. The architecture was carefully planned by humans, but some straightforward details were implemented by AI. All AI contributions were reviewed by a human, and the output of the software was compared with human computed output on a large number of real world data sets uploaded to Zenodo. After downloading the paper's BibTeX file from Zotero, Gemini 3 Flash was used to refine and format this initial file; all references were compiled, organized, and verified by the authors.  
