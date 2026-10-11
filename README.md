# collect_ms_data.py</br>
  Vibe coded collector for MS CVE data, output directly ingestable by Grafana tables.</br>
  Have not done any thorouhg testing, use at you own discression.</br>
  </br>
# collect_oracle_data.py</br>
  Vibe coded collector for Oracle CVE data, output directly ingestable by Grafana tables.</br>
  Not thoroughly tested either, also produces a lot of unnecessary data I didn't find the interest in cleaning up.</br>
</br>
# mbox-parser.py</br>
  Just an awful mess atm. Something I whipped up while doing some investigation into mbox content and needed someting to parse them with, without using an actual client.</br>
</br>
  The HTML-exctract depends on:</br>
    bootstrap.min.css</br>
    jquery-3.2.1.slim.min.js</br>
    bootstrap.min.js</br>
</br>
At the moment creates everything inside the folder it was run from into the output folder, under the mailbox name. Maybe one day I have time to organize the code, do proper division to functions and classes etc. but not today.
