---
layout: archive
title: "CV"
permalink: /cv/
author_profile: true
redirect_from:
  - /resume
---

{% include base_path %}

Education
======

* B.A. in **Psychology**, Federal University of Health Sciences of Porto Alegre, 2021
* M.A. in **Psychology**, Federal University of Rio Grande do Sul, 2022
* Ph.D. in **Literary, Cultural, and Linguistic Studies**, University of Miami, 2027 (expected)

Work experience
======
* Summer 2015: Research Assistant
  * Github University
  * Duties included: Tagging issues
  * Supervisor: Professor Git

* Fall 2015: Research Assistant
  * Github University
  * Duties included: Merging pull requests
  * Supervisor: Professor Hub

Publications
======

{% include_relative publications.md %}
  
Talks
======

{% include_relative talks.md %}
  
Teaching
======

{% capture teaching_content %}{% include_relative teaching.md %}{% endcapture %}
{{ teaching_content | remove_first: "---\n" | markdownify }}
  
