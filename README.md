AsciiEXM
========

Written by Christopher R. Harty 2009

Brick layer backup of Microsoft Exchange 2K3 or 2K7 mailboxes.  

Requirements:
<ul>
  <li>WINXP SP2 and is a domain member</li>
  <li>VB6</li>
  <li>Microsoft Exmerge API</li>
  <li>Logged in as Exchange Admin</li>
</ul>

Functionality:
<ul>
  <li>UI will allow for choosing of any Exchange Server found within its domain.</li>
  <li>Choose a Group(OU) within the specified Server</li>
  <li>Drill down and then choose a Mailbox Database from selected Group</li>
  --(preview of all accounts along with # of message and mailbox size)</li>
  <li>Multi-select user accounts to backup</li>
  --(displays cumulative disk space to be used and accounts being backedup)</li>
  <li>Progress UI displaying logged results for a batch of mailboxes being backed up</li>
</ul>

Options: 
<ul>
  <li>Compression</li>
  <li>Encryption</li>
  <li>2GB or monthly partitioning</li>
  <li>Threading controls</li>
</ul>

Settings: 
<ul>
  <li>Set config, logs, and target mailbox file paths(PST).  Sub folders created, if they don't exist</li>
  <li>Mailbox.txt Template Settings</li>
  <li>Date and time ranges for batch</li>
  <li>Granular checkboxes for User Data, Dumpster, Folder Data, and Folder Rules
</ul>
