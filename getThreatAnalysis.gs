/**
 * AdminPulse Threat Analysis Extension
 * Compiles all misconfigurations and categorizes them by threat type
 */

function runThreatAnalysis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  try {
    // Get or create the Threat Analysis sheet
    let threatSheet = ss.getSheetByName('Threat Analysis');
    if (!threatSheet) {
      threatSheet = ss.insertSheet('Threat Analysis');
    } else {
      threatSheet.clear();
    }
    
    // Collect all findings
    let allFindings = [];
    allFindings = allFindings.concat(analyzeUsersSheet(ss));
    allFindings = allFindings.concat(analyzeOAuthSheet(ss));
    allFindings = allFindings.concat(analyzeGroupsSheet(ss));
    allFindings = allFindings.concat(analyzeSharedDrivesSheet(ss));
    allFindings = allFindings.concat(analyzeAppPasswordsSheet(ss)); 
    allFindings = allFindings.concat(analyzeAliasesSheet(ss));
    allFindings = allFindings.concat(analyzeGroupMembersSheet(ss));
    allFindings = allFindings.concat(analyzeAdditionalServicesSheet(ss));
    
    // Sort by threat category, then risk level
    const riskOrder = { 'Critical': 0, 'High': 1, 'Medium': 2, 'Low': 3 };
    allFindings.sort((a, b) => {
      if (a[0] !== b[0]) return a[0].localeCompare(b[0]);
      return riskOrder[a[3]] - riskOrder[b[3]];
    });
    
    // Group findings by threat category
    const groupedFindings = {};
    allFindings.forEach(finding => {
      const category = finding[0];
      if (!groupedFindings[category]) {
        groupedFindings[category] = [];
      }
      groupedFindings[category].push(finding);
    });
    
    // Write findings grouped by threat category
    let currentRow = 1;
    
    // Title
    threatSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('ADMINPULSE THREAT ANALYSIS REPORT')
      .setFontSize(16)
      .setFontWeight('bold')
      .setBackground('#fc3165')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    currentRow += 2;

    // Data Sources information
    threatSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('This analysis uses data from the following AdminPulse reports:')
      .setFontWeight('bold')
      .setFontSize(11);
    currentRow++;
    
    threatSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('Users, OAuth Tokens, Group Settings, Shared Drives, App Passwords, Aliases, Group Members and Additional Services')
      .setFontStyle('italic')
      .setFontColor('#555555')
      .setWrap(true);
    currentRow++;
    
    threatSheet.getRange(currentRow, 1, 1, 5).merge()
      .setValue('Findings are categorized by threat type to help prioritize remediation based on potential security impact.')
      .setFontStyle('italic')
      .setFontColor('#555555')
      .setWrap(true);
    currentRow += 2;
    
    // Summary stats
    threatSheet.getRange(currentRow, 1).setValue('Total Findings:').setFontWeight('bold');
    threatSheet.getRange(currentRow, 2).setValue(allFindings.length).setFontWeight('bold');
    currentRow++;
    

    currentRow += 2;
    
    // Define threat category order and descriptions
    const categoryOrder = [
      'Account Compromise',
      'Data Exfiltration Routes',
      'Potential Data Leaks',
      'Lateral Movement Points',
      'Compliance/Audit Gaps'
    ];
    
    const categoryDescriptions = {
      'Account Compromise': 'Settings that make account compromises more likely or impactful',
      'Data Exfiltration Routes': 'Configurations that enable data exfiltration if an account is compromised',
      'Potential Data Leaks': 'Settings that could lead to accidental exposure of sensitive information',
      'Lateral Movement Points': 'Configurations that help attackers move between accounts or escalate privileges',
      'Compliance/Audit Gaps': 'Missing controls for monitoring, logging, and compliance requirements'
    };
    
    // Write each threat category section
    categoryOrder.forEach(category => {
      if (!groupedFindings[category] || groupedFindings[category].length === 0) return;
      
      const findings = groupedFindings[category];
      
      // Category header
      threatSheet.getRange(currentRow, 1, 1, 5).merge()
        .setValue('━━━ ' + category.toUpperCase() + ' ━━━')
        .setFontSize(14)
        .setFontWeight('bold')
        .setBackground('#000433')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center');
      currentRow++;
      
      // Category description
      threatSheet.getRange(currentRow, 1, 1, 5).merge()
        .setValue(categoryDescriptions[category])
        .setFontStyle('italic')
        .setFontColor('#7f8c8d')
        .setWrap(true);
      currentRow++;
      
      // Finding count
      threatSheet.getRange(currentRow, 1, 1, 2)
        .setValues([['Risks found:', findings.length]])
        .setFontWeight('bold');
      currentRow += 2;
      
      // Column headers for this section
      threatSheet.getRange(currentRow, 1, 1, 4)
        .setValues([['Risk', 'Finding Type', 'Affected Item', 'Remediation']])
        .setBackground('#95a5a6')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
      currentRow++;
      
      // Write findings for this category
      findings.forEach(finding => {
        const [cat, findingType, affectedItem, risk, remediation] = finding;
        
        threatSheet.getRange(currentRow, 1, 1, 4).setValues([[
          risk,
          findingType,
          affectedItem,
          remediation
        ]]);
        
        // Color code risk level
        let riskColor = '#ffffff';
        switch(risk) {
          case 'Critical': riskColor = '#e74c3c'; break;
          case 'High': riskColor = '#e67e22'; break;
          case 'Medium': riskColor = '#f39c12'; break;
          case 'Low': riskColor = '#2ecc71'; break;
        }
        threatSheet.getRange(currentRow, 1).setBackground(riskColor).setFontColor('#ffffff').setFontWeight('bold');
        
        currentRow++;
      });
      
      currentRow += 2; // Space between categories
    });
    
    // Auto-resize columns
    threatSheet.setColumnWidth(1, 100);  // Risk
    threatSheet.setColumnWidth(2, 300);  // Finding Type
    threatSheet.setColumnWidth(3, 350);  // Affected Item
    threatSheet.setColumnWidth(4, 450);  // Remediation
    
    // Freeze header rows
    threatSheet.setFrozenRows(1);
    
    Browser.msgBox('Threat Analysis Complete', 
                   'Found ' + allFindings.length + ' security findings across ' + 
                   Object.keys(groupedFindings).length + ' threat categories. ' +
                   'Check the "Threat Analysis" sheet for the complete report.',
                   Browser.Buttons.OK);
    
  } catch (error) {
    Browser.msgBox('Error', 'An error occurred: ' + error.toString(), Browser.Buttons.OK);
    Logger.log('Error: ' + error.toString());
  }
}

function analyzeUsersSheet(ss) {
  const findings = [];
  const sheet = ss.getSheetByName('Users');
  if (!sheet) return findings;
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return findings;
    
    const headers = data[0];
    const emailIdx = headers.indexOf('Email');
    const superAdminIdx = headers.indexOf('Super Admin');
    const delegatedAdminIdx = headers.indexOf('Delegated Admin');
    const twoSVIdx = headers.indexOf('Enrolled in 2SV');
    const suspendedIdx = headers.indexOf('Suspended');
    const archivedIdx = headers.indexOf('Archived');
    const lastLoginIdx = headers.indexOf('Last Login Time');
    
    // Get the report generation timestamp from Workspace Security Checklist sheet
    let reportDate = new Date(); // Default to now if timestamp not found
    try {
      const checklistSheet = ss.getSheetByName('Workspace Security Checklist');
      if (checklistSheet) {
        const timestampCell = checklistSheet.getRange('F1').getValue();
        if (timestampCell) {
          const timestampStr = timestampCell.toString();
          // Extract date from "Last policy inventory completed at: [date]" format
          const match = timestampStr.match(/(\d{1,2}\/\d{1,2}\/\d{4}[,\s]+\d{1,2}:\d{2}:\d{2}\s*(?:AM|PM)?)/i);
          if (match) {
            reportDate = new Date(match[1]);
          }
        }
      }
    } catch (e) {
      Logger.log('Could not read report timestamp, using current date: ' + e);
    }
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[emailIdx]) continue;
      
      const email = row[emailIdx].toString();
      const isSuperAdmin = row[superAdminIdx] === true || row[superAdminIdx] === 'TRUE';
      const isDelegatedAdmin = row[delegatedAdminIdx] === true || row[delegatedAdminIdx] === 'TRUE';
      const has2SV = row[twoSVIdx] === true || row[twoSVIdx] === 'TRUE';
      const isSuspended = row[suspendedIdx] === true || row[suspendedIdx] === 'TRUE';
      const isArchived = row[archivedIdx] === true || row[archivedIdx] === 'TRUE';
      const lastLoginRaw = lastLoginIdx >= 0 ? row[lastLoginIdx] : null;

      if (isSuspended || isArchived) continue;

      // 2FA findings
      if (isSuperAdmin && !has2SV) {
        findings.push([
          'Account Compromise',
          'Super Admin without 2-Step Verification',
          email,
          'Critical',
          'Enforce 2SV immediately. Super admins have full domain control.'
        ]);
      }

      if (isDelegatedAdmin && !has2SV && !isSuperAdmin) {
        findings.push([
          'Account Compromise',
          'Delegated Admin without 2-Step Verification',
          email,
          'Critical',
          'Enforce 2SV for all privileged accounts.'
        ]);
      }

      if (!isSuperAdmin && !isDelegatedAdmin && !has2SV) {
        findings.push([
          'Account Compromise',
          'Active user without 2-Step Verification',
          email,
          'High',
          'Implement organization-wide 2SV enforcement.'
        ]);
      }

      // Inactivity check using report generation date
      let daysSinceLogin = null;

      if (lastLoginRaw) {
        try {
          const lastLogin = new Date(lastLoginRaw);
          daysSinceLogin = Math.floor((reportDate - lastLogin) / (1000 * 60 * 60 * 24));
        } catch (e) {
          daysSinceLogin = null;
        }
      }

      const isInactive = daysSinceLogin === null || daysSinceLogin >= 90;

      if (isInactive) {
        const risk = (isSuperAdmin || isDelegatedAdmin) ? 'High' : 'Medium';
        const daysText = daysSinceLogin === null ? 'never logged in' : 'no login for ' + daysSinceLogin + ' days';

        findings.push([
          'Compliance/Audit Gaps',
          'Inactive account (' + daysText + ')',
          email,
          risk,
          'Disable, suspend, or remove accounts inactive for 90+ days to reduce attack surface and meet audit requirements.'
        ]);
      }
    }
  } catch (error) {
    Logger.log('Error analyzing Users sheet: ' + error);
  }
  
  return findings;
}

function analyzeOAuthSheet(ss) {
  const findings = [];
  const sheet = ss.getSheetByName('OAuth Tokens');
  if (!sheet) return findings;
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return findings;
    
    const headers = data[0];
    const emailIdx = headers.indexOf('User Email');
    const appNameIdx = headers.indexOf('Application Name');
    const clientIdIdx = headers.indexOf('Client ID');
    const scopesIdx = headers.indexOf('Granted Scopes');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[emailIdx]) continue;
      
      const email = row[emailIdx].toString();
      const appName = row[appNameIdx] ? row[appNameIdx].toString() : 'Unknown App';
      const clientId = row[clientIdIdx] ? row[clientIdIdx].toString().substring(0, 20) + '...' : '';
      const scopes = row[scopesIdx] ? row[scopesIdx].toString().toLowerCase() : '';
      
      // Critical: Admin SDK access
      if (scopes.includes('admin.directory')) {
        findings.push([
          'Data Exfiltration Routes',
          'Third-party app with Admin SDK access',
          appName + ' → ' + email,
          'Critical',
          'Verify app legitimacy immediately. Admin SDK allows complete domain control including user/group management.'
        ]);
      }
      
      // High: Full Drive access
      if (scopes.includes('drive') && !scopes.includes('drive.file')) {
        findings.push([
          'Data Exfiltration Routes',
          'Third-party app with full Google Drive access',
          appName + ' → ' + email,
          'High',
          'Review if full Drive access is necessary. App can read/modify all Drive files. Consider limiting to drive.file scope.'
        ]);
      }
      
      // High: Full Gmail access
      if (scopes.includes('mail.google.com')) {
        findings.push([
          'Data Exfiltration Routes',
          'Third-party app with full Gmail access',
          appName + ' → ' + email,
          'High',
          'Verify necessity of full Gmail access. App can read, send, and delete all emails.'
        ]);
      }
      
      // Medium: Calendar access
      if (scopes.includes('calendar')) {
        findings.push([
          'Potential Data Leaks',
          'Third-party app with Calendar access',
          appName + ' → ' + email,
          'Low',
          'Review if calendar access is still needed. May expose meeting details and schedules.'
        ]);
      }
      
      // Medium: Contacts access
      if (scopes.includes('contacts')) {
        findings.push([
          'Potential Data Leaks',
          'Third-party app with Contacts access',
          appName + ' → ' + email,
          'Low',
          'Verify contacts access is required. May expose organizational directory information.'
        ]);
      }
    }
  } catch (error) {
    Logger.log('Error analyzing OAuth sheet: ' + error);
  }
  
  return findings;
}

function analyzeGroupsSheet(ss) {
  const findings = [];
  const sheet = ss.getSheetByName('Group Settings');
  if (!sheet) return findings;
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return findings;
    
    const headers = data[0];
    const nameIdx = headers.indexOf('name');
    const emailIdx = headers.indexOf('email');
    const allowExternalIdx = headers.indexOf('allowExternalMembers');
    const whoCanPostIdx = headers.indexOf('whoCanPostMessage');
    const whoCanViewMemberIdx = headers.indexOf('whoCanViewMembership');
    const whoCanDiscoverIdx = headers.indexOf('whoCanDiscoverGroup');
    const allowWebPostingIdx = headers.indexOf('allowWebPosting');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[emailIdx]) continue;
      
      const groupName = row[nameIdx] ? row[nameIdx].toString() : 'Unknown';
      const groupEmail = row[emailIdx].toString();
      const allowExternal = row[allowExternalIdx];
      const whoCanPost = row[whoCanPostIdx] ? row[whoCanPostIdx].toString() : '';
      const whoCanView = row[whoCanViewMemberIdx] ? row[whoCanViewMemberIdx].toString() : '';
      const whoCanDiscover = row[whoCanDiscoverIdx] ? row[whoCanDiscoverIdx].toString() : '';
      const allowWebPosting = row[allowWebPostingIdx];
      
      // High: Anyone can post (internet-wide)
      if (whoCanPost === 'ANYONE_CAN_POST') {
        findings.push([
          'Potential Data Leaks',
          'Group allows anyone on internet to post messages',
          groupName + ' (' + groupEmail + ')',
          'High',
          'Restrict posting permissions immediately. This allows spam and potential data injection attacks.'
        ]);
      }
      
      // Medium: External members allowed
      if (allowExternal === true || allowExternal === 'true' || allowExternal === 'TRUE') {
        findings.push([
          'Potential Data Leaks',
          'Group allows external (non-domain) members',
          groupName + ' (' + groupEmail + ')',
          'Medium',
          'Review if external membership is business-necessary. External members can see group content.'
        ]);
      }
      
      // Medium: All domain users can post
      if (whoCanPost === 'ALL_IN_DOMAIN_CAN_POST') {
        findings.push([
          'Lateral Movement Points',
          'All domain users can post to group',
          groupName + ' (' + groupEmail + ')',
          'Low',
          'Consider restricting to members-only posting if group handles sensitive information.'
        ]);
      }
      
      // Low: All domain can view membership
      if (whoCanView === 'ALL_IN_DOMAIN_CAN_VIEW') {
        findings.push([
          'Lateral Movement Points',
          'All domain users can view group membership',
          groupName + ' (' + groupEmail + ')',
          'Low',
          'Restrict membership visibility if this group represents sensitive organizational structure.'
        ]);
      }
      
      // Medium: Publicly discoverable
      if (whoCanDiscover === 'ALL_IN_DOMAIN_CAN_DISCOVER' || whoCanDiscover === 'ANYONE_CAN_DISCOVER') {
        findings.push([
          'Compliance/Audit Gaps',
          'Group is publicly discoverable',
          groupName + ' (' + groupEmail + ')',
          'Medium',
          'Limit discoverability if group name or purpose is sensitive.'
        ]);
      }
      
      // Medium: Web posting enabled
      if (allowWebPosting === true || allowWebPosting === 'true' || allowWebPosting === 'TRUE') {
        findings.push([
          'Potential Data Leaks',
          'Group allows posting via web interface',
          groupName + ' (' + groupEmail + ')',
          'Low',
          'Disable web posting if not needed to prevent unauthorized content publication.'
        ]);
      }
    }
  } catch (error) {
    Logger.log('Error analyzing Groups sheet: ' + error);
  }
  
  return findings;
}

function analyzeSharedDrivesSheet(ss) {
  const findings = [];
  const sheet = ss.getSheetByName('Shared Drives');
  if (!sheet) return findings;
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return findings;
    
    const headers = data[0];
    const nameIdx = headers.indexOf('Name');
    const copyRequiresWriterIdx = headers.indexOf('Copy Requires Writer Permission');
    const domainUsersOnlyIdx = headers.indexOf('Domain Users Only');
    const driveMembersOnlyIdx = headers.indexOf('Drive Members Only');
    const adminManagedIdx = headers.indexOf('Admin Managed Restrictions');
    const sharingFoldersIdx = headers.indexOf('Sharing Folders Requires Organizer Permission');
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[nameIdx]) continue;
      
      const driveName = row[nameIdx].toString();
      const copyRequiresWriter = row[copyRequiresWriterIdx];
      const domainUsersOnly = row[domainUsersOnlyIdx];
      const driveMembersOnly = row[driveMembersOnlyIdx];
      const adminManaged = row[adminManagedIdx];
      const sharingFoldersRequiresOrg = row[sharingFoldersIdx];
      
      // 1. Copy/Download restrictions
      if (copyRequiresWriter === false || copyRequiresWriter === 'FALSE' || copyRequiresWriter === 'false') {
        findings.push([
          'Data Exfiltration Routes',
          'Shared Drive allows anyone to copy/download files',
          driveName,
          'High',
          'Enable "Viewers and commenters cannot download" restriction to prevent unauthorized data exfiltration.'
        ]);
      }
      
      // 2. Domain users only
      if (domainUsersOnly === false || domainUsersOnly === 'FALSE' || domainUsersOnly === 'false') {
        findings.push([
          'Potential Data Leaks',
          'Shared Drive allows external (non-domain) user access',
          driveName,
          'High',
          'Restrict access to domain users only unless external collaboration is explicitly required for business purposes.'
        ]);
      }
      
      // 3. Drive members only
      if (driveMembersOnly === false || driveMembersOnly === 'FALSE' || driveMembersOnly === 'false') {
        findings.push([
          'Potential Data Leaks',
          'Shared Drive allows non-members to access shared items',
          driveName,
          'Medium',
          'Enable "Only users explicitly granted access" to prevent link-based sharing bypass.'
        ]);
      }
      
      // 4. Admin managed restrictions
      if (adminManaged === false || adminManaged === 'FALSE' || adminManaged === 'false') {
        findings.push([
          'Compliance/Audit Gaps',
          'Shared Drive restrictions can be modified by managers',
          driveName,
          'High',
          'Enable admin-managed restrictions to prevent drive managers from weakening security controls.'
        ]);
      }
      
      // 5. Sharing folders requires organizer permission
      if (sharingFoldersRequiresOrg === false || sharingFoldersRequiresOrg === 'FALSE' || sharingFoldersRequiresOrg === 'false') {
        findings.push([
          'Lateral Movement Points',
          'Shared Drive allows any member to share folders',
          driveName,
          'Medium',
          'Require organizer permission for folder sharing to maintain centralized control over access grants.'
        ]);
      }
    }
  } catch (error) {
    Logger.log('Error analyzing Shared Drives sheet: ' + error);
  }
  
  return findings;
}

function analyzeAliasesSheet(ss) {
  const findings = [];
  const sheet = ss.getSheetByName('Aliases');
  if (!sheet) return findings;
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return findings;
    
    const headers = data[0];
    const emailIdx = headers.indexOf('primaryEmail');
    const aliasIdx = headers.indexOf('alias');
    
    const aliasesByUser = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[emailIdx]) continue;
      
      const email = row[emailIdx].toString();
      const alias = row[aliasIdx] ? row[aliasIdx].toString() : '';
      
      if (!aliasesByUser[email]) {
        aliasesByUser[email] = [];
      }
      if (alias) {
        aliasesByUser[email].push(alias);
      }
    }
    
    for (const email in aliasesByUser) {
      const aliases = aliasesByUser[email];
      if (aliases.length >= 5) {
        findings.push([
          'Lateral Movement Points',
          'User has excessive email aliases (' + aliases.length + ')',
          email,
          'Low',
          'Review if all aliases are necessary. Excessive aliases can obscure audit trails and phishing attempts.'
        ]);
      }
    }
  } catch (error) {
    Logger.log('Error analyzing Aliases sheet: ' + error);
  }
  
  return findings;
}

function analyzeAppPasswordsSheet(ss) {
  const findings = [];
  const sheet = ss.getSheetByName('App Passwords');
  if (!sheet) return findings;
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return findings;
    
    const headers = data[0];
    const nameIdx = headers.indexOf('Name');
    const userIdx = headers.indexOf('User');
    
    const appPasswordsByUser = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[userIdx]) continue;
      
      const user = row[userIdx].toString();
      const appName = row[nameIdx] ? row[nameIdx].toString() : 'Unknown App';
      
      if (!appPasswordsByUser[user]) {
        appPasswordsByUser[user] = [];
      }
      appPasswordsByUser[user].push(appName);
    }
    
    // Report on users with app passwords
    for (const user in appPasswordsByUser) {
      const apps = appPasswordsByUser[user];
      findings.push([
        'Account Compromise',
        'User has app-specific password(s) that bypass 2FA',
        user + ' (' + apps.length + ' app password' + (apps.length > 1 ? 's' : '') + ')',
        'Medium',
        'App passwords bypass 2-Step Verification. Migrate to OAuth2 apps where possible and revoke unused app passwords.'
      ]);
    }
    
  } catch (error) {
    Logger.log('Error analyzing App Passwords sheet: ' + error);
  }
  
  return findings;
}

function analyzeGroupMembersSheet(ss) {
  const findings = [];
  const sheet = ss.getSheetByName('Group Members');
  if (!sheet) return findings;
  
  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return findings;
    
    const headers = data[0];
    const groupEmailIdx = headers.indexOf('groupEmail');
    const memberEmailIdx = headers.indexOf('email');
    const roleIdx = headers.indexOf('role');
    const typeIdx = headers.indexOf('type');
    
    // Track external members in groups
    const externalMembers = {};
    const ownersByGroup = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[groupEmailIdx] || !row[memberEmailIdx]) continue;
      
      const groupEmail = row[groupEmailIdx].toString();
      const memberEmail = row[memberEmailIdx].toString();
      const role = row[roleIdx] ? row[roleIdx].toString() : '';
      const type = row[typeIdx] ? row[typeIdx].toString() : '';
      
      // Check for external members based on the 'type' column
      if (type === 'EXTERNAL') {
        if (!externalMembers[groupEmail]) {
          externalMembers[groupEmail] = [];
        }
        externalMembers[groupEmail].push(memberEmail);
      }
      
      // Count owners per group
      if (role === 'OWNER') {
        if (!ownersByGroup[groupEmail]) {
          ownersByGroup[groupEmail] = [];
        }
        ownersByGroup[groupEmail].push(memberEmail);
      }
    }
    
    // Report groups with external members
    for (const group in externalMembers) {
      const members = externalMembers[group];
      findings.push([
        'Potential Data Leaks',
        'Group has external members (' + members.length + ')',
        group + ' → ' + members.slice(0, 3).join(', ') + (members.length > 3 ? '...' : ''),
        'Medium',
        'Review if external members should have access to this group\'s content.'
      ]);
    }
    
    // Report groups with too many owners
    for (const group in ownersByGroup) {
      const owners = ownersByGroup[group];
      if (owners.length >= 5) {
        findings.push([
          'Lateral Movement Points',
          'Group has excessive owners (' + owners.length + ')',
          group,
          'Low',
          'Limit group ownership to necessary personnel. Too many owners increases risk of misconfiguration.'
        ]);
      }
    }
  } catch (error) {
    Logger.log('Error analyzing Group Members sheet: ' + error);
  }
  
  return findings;
}

function analyzeAdditionalServicesSheet(ss) {
  const findings = [];
  const sheet = ss.getSheetByName('Additional Services');
  if (!sheet) return findings;

  try {
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return findings;

    const headers = data[0].map(h => h.toString().toLowerCase());
    const serviceIdx = headers.indexOf('service') !== -1 ? headers.indexOf('service') : 0;
    const ouIdx = headers.indexOf('ou / group') !== -1 ? headers.indexOf('ou / group') : 1;
    const statusIdx = headers.indexOf('status') !== -1 ? headers.indexOf('status') : 2;

    const SERVICE_RISK_MAP = {
      'ad_manager': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Ad Manager controls advertising inventory. Restrict to marketing team.' },
      'ads': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Google Ads may access customer data for targeting. Review data sharing settings.' },
      'adsense': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'AdSense displays ads. Monitor for inappropriate content and brand safety.' },
      'ai_studio': { category: 'Account Compromise', risk: 'Medium', reason: 'AI Studio processes sensitive prompts. Review data handling policies.' },
      'alerts': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Security alerts notify of suspicious activity. Ensure proper monitoring.' },
      'analytics': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Analytics may contain PII. Ensure GDPR/CCPA compliance.' },
      'applied_digital_skills': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Educational platform. May be unnecessary in corporate environments.' },
      'appsheet': { category: 'Account Compromise', risk: 'High', reason: 'AppSheet creates custom apps with data access. Review app permissions carefully.' },
      'arts_and_culture': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Cultural exploration. Personal service; typically unnecessary in business.' },
      'assignments': { category: 'Potential Data Leaks', risk: 'Medium', reason: 'Assignment data includes student work. Verify FERPA/COPPA compliance.' },
      'blogger': { category: 'Potential Data Leaks', risk: 'Low', reason: 'Blogger posts are public by default. Ensure user awareness.' },
      'bookmarks': { category: 'Data Exfiltration Routes', risk: 'Low', reason: 'Bookmark sync may reveal internal URLs. Consider disabling.' },
      'books': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Google Books for personal use. Disable unless needed for library/research.' },
      'brand_accounts': { category: 'Lateral Movement Points', risk: 'Low', reason: 'Brand accounts enable multi-user access. Audit membership regularly.' },
      'calendar': { category: 'Lateral Movement Points', risk: 'Medium', reason: 'Calendar provides a blueprint of organizational hierarchy and sensitive activities. Restrict internal sharing options properly.' },
      'campaign_manager': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Campaign Manager tracks marketing. Ensure privacy compliance.' },
      'chat': { category: 'Lateral Movement Points', risk: 'Medium', reason: 'Chat serves as a high-trust lateral movement vector that attackers abuse to pivot across an organization. Disable unless required.' },
      'chrome_canvas': { category: 'Potential Data Leaks', risk: 'Low', reason: 'Canvas may contain sensitive diagrams. Monitor sharing settings.' },
      'chrome_cursive': { category: 'Potential Data Leaks', risk: 'Low', reason: 'Handwriting synced to cloud. Review if notes are sensitive.' },
      'chrome_remote_desktop': { category: 'Account Compromise', risk: 'High', reason: 'Remote desktop creates persistence for attackers. Disable unless essential.' },
      'chrome_sync': { category: 'Data Exfiltration Routes', risk: 'High', reason: 'Browser sync transmits passwords and history. Restrict to managed devices.' },
      'chrome_web_store': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Chrome extension installation. Restrict to approved extensions.' },
      'classroom': { category: 'Potential Data Leaks', risk: 'Medium', reason: 'Classroom contains student data. Ensure FERPA/COPPA compliance.' },
      'cloud': { category: 'Account Compromise', risk: 'High', reason: 'Google Cloud grants infrastructure access. Ensure proper IAM controls.' },
      'colab': { category: 'Account Compromise', risk: 'High', reason: 'Colab executes code. Attackers could run malicious scripts. Restrict access.' },
      'cs_first': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Computer science education. Disable if not running educational programs.' },
      'data_studio': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Data Studio aggregates business metrics. Review sharing settings.' },
      'developers': { category: 'Account Compromise', risk: 'High', reason: 'Developer tools grant API access. Restrict to authorized teams.' },
      'domains': { category: 'Lateral Movement Points', risk: 'Medium', reason: 'Domain management controls identity. Restrict to authorized admins.' },
      'drive_and_docs': { category: 'Data Exfiltration Routes', risk: 'Medium', reason: 'Drive contains organizational documents. Ensure DLP and access controls.' },
      'early_access_apps': { category: 'Account Compromise', risk: 'Medium', reason: 'Early access may have vulnerabilities. Evaluate before deployment.' },
      'earth': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Google Earth for mapping. Disable if not needed for GIS work.' },
      'enterprise_service_restrictions': { category: 'Lateral Movement Points', risk: 'Low', reason: 'Service restrictions control features. Misconfigurations bypass controls.' },
      'experimental_apps': { category: 'Account Compromise', risk: 'Medium', reason: 'Experimental apps lack mature security. Use only in non-production.' },
      'feedburner': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'RSS feed management. Legacy tool; review if still used.' },
      'fifi': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Internal Google service. Review with support if uncertain.' },
      'gemini_app': { category: 'Account Compromise', risk: 'Medium', reason: 'Gemini processes organizational queries. Verify data compliance.' },
      'gmail': { category: 'Lateral Movement Points', risk: 'High', reason: 'Gmail delegation can be abused by attackers to pivot to other mailboxes. Review Gmail delegation if implemented.' },
      'groups': { category: 'Potential Data Leaks', risk: 'Medium', reason: 'Groups share information. Review membership and external access policies.' },
      'groups_for_business': { category: 'Potential Data Leaks', risk: 'Medium', reason: 'Business groups may contain sensitive discussions. Audit external access.' },
      'jamboard': { category: 'Potential Data Leaks', risk: 'Low', reason: 'Jamboard contains brainstorming sessions. Review sharing permissions.' },
      'keep': { category: 'Potential Data Leaks', risk: 'Low', reason: 'Keep notes may be sensitive. Review sharing to prevent external exposure.' },
      'location_history': { category: 'Compliance/Audit Gaps', risk: 'Medium', reason: 'Location tracking raises privacy concerns. Disable unless required.' },
      'managed_play': { category: 'Lateral Movement Points', risk: 'Low', reason: 'Controls app installation. Review approved list regularly.' },
      'maps': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Google Maps may track location. Review if business requires it.' },
      'material_gallery': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Design resource library. Relevant for design teams only.' },
      'meet': { category: 'Potential Data Leaks', risk: 'Medium', reason: 'Meet recordings contain sensitive discussions. Ensure access controls.' },
      'merchant_center': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Merchant Center handles payments. Ensure PCI DSS compliance.' },
      'messages': { category: 'Potential Data Leaks', risk: 'Low', reason: 'Messages syncs SMS/MMS. Verify if business communications should backup.' },
      'my_business': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Business profile is public. Monitor for unauthorized changes.' },
      'my_maps': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Custom map creation. Disable if not used for logistics.' },
      'nest': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Smart home management. Personal service; disable in business.' },
      'news': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Google News aggregator. Disable unless business-critical.' },
      'notebooklm': { category: 'Account Compromise', risk: 'Medium', reason: 'NotebookLM processes documents. Verify no sensitive document sharing.' },
      'partner_dash': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Partner dashboard. Limit to partnership/business development.' },
      'pay': { category: 'Compliance/Audit Gaps', risk: 'Medium', reason: 'Google Pay enables payments. Disable unless required; ensure financial controls.' },
      'photos': { category: 'Data Exfiltration Routes', risk: 'Low', reason: 'Photos may contain screenshots of sensitive info. Review if appropriate.' },
      'pinpoint': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Document analysis for journalists. Disable unless investigative research.' },
      'play': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Google Play downloads. Review if app installation should be restricted.' },
      'play_books_partner_center': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Book publishing. Disable unless publishing digital books.' },
      'play_console': { category: 'Lateral Movement Points', risk: 'Low', reason: 'App development platform. Limit to development team.' },
      'public_data': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Access to public datasets. Disable unless needed for data research.' },
      'question_hub': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Q&A for content creators. Disable unless content/marketing programs.' },
      'read_along': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Reading assistance for children. Disable unless K-12 programs.' },
      'scholar_profiles': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Academic publications. Relevant for research institutions only.' },
      'search_ads_360': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Advertising management. Limit to marketing team.' },
      'search_and_assistant': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Google Assistant may process voice commands with sensitive info.' },
      'search_console': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Website analytics. Limit to marketing/SEO staff.' },
      'sites': { category: 'Potential Data Leaks', risk: 'Medium', reason: 'Sites can be shared publicly. Audit for accidental exposure.' },
      'socratic': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Homework help. Unnecessary in corporate; disable if not needed.' },
      'takeout': { category: 'Data Exfiltration Routes', risk: 'High', reason: 'Takeout enables bulk export. Attackers can exfiltrate all data. Disable unless required.' },
      'tasks': { category: 'Potential Data Leaks', risk: 'Low', reason: 'Tasks may reveal project details. Ensure not shared externally.' },
      'third_party_app_backups': { category: 'Data Exfiltration Routes', risk: 'High', reason: 'Third-party backups copy all data. Vet providers thoroughly.' },
      'translate': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Translation processes text via Google. Ensure users understand.' },
      'trips': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Travel itinerary. Personal service; disable unless travel management.' },
      'vault': { category: 'Data Exfiltration Routes', risk: 'Critical', reason: 'Vault houses the historical record for an organization as a centralized and searchable interface. Review access control implementation.' },
      'voice': { category: 'Compliance/Audit Gaps', risk: 'Low', reason: 'Google Voice telephony. Review if VoIP should be enabled.' },
      'youtube': { category: 'Potential Data Leaks', risk: 'Low', reason: 'YouTube may expose training videos or recordings. Monitor channel content.' }
    };

    const seen = new Set();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[serviceIdx] || !row[statusIdx]) continue;

      const serviceRaw = row[serviceIdx].toString().trim();
      const serviceKey = serviceRaw.toLowerCase().replace(/[_\s]+/g, '_');
      if (seen.has(serviceKey)) continue;

      const statusValue = row[statusIdx].toString().toLowerCase();
      const enabled = statusValue.includes('enabled') || statusValue === 'true';

      if (!enabled) continue;

      seen.add(serviceKey);

      const scope = ouIdx >= 0 && row[ouIdx] ? row[ouIdx].toString() : '/';

      let matched = false;

      for (const serviceName in SERVICE_RISK_MAP) {
        if (serviceKey.includes(serviceName.replace(/_/g, '')) || serviceKey === serviceName) {
          const mapping = SERVICE_RISK_MAP[serviceName];

          findings.push([
            mapping.category,
            'Workspace service enabled: ' + serviceRaw,
            'Scope: ' + scope,
            mapping.risk,
            mapping.reason
          ]);

          matched = true;
          break;
        }
      }

      if (!matched) {
        findings.push([
          'Compliance/Audit Gaps',
          'Enabled service requires review: ' + serviceRaw,
          'Scope: ' + scope,
          'Low',
          'Validate business need. Disable unused services to reduce attack surface.'
        ]);
      }
    }
  } catch (error) {
    Logger.log('Error analyzing Additional Services: ' + error);
  }

  return findings;
}
