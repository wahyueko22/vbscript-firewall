'  This VBScript file includes sample code that enumerates
'  Windows Firewall rules using the Microsoft Windows Firewall APIs.


Option Explicit

Dim CurrentProfiles
Dim InterfaceArray
Dim LowerBound
Dim UpperBound
Dim iterate
Dim ruleIterate
Dim rule

' Profile Type
Const NET_FW_PROFILE2_DOMAIN = 1
Const NET_FW_PROFILE2_PRIVATE = 2
Const NET_FW_PROFILE2_PUBLIC = 4

' Protocol
Const NET_FW_IP_PROTOCOL_TCP = 6
Const NET_FW_IP_PROTOCOL_UDP = 17
Const NET_FW_IP_PROTOCOL_ICMPv4 = 1
Const NET_FW_IP_PROTOCOL_ICMPv6 = 58

' Direction
Const NET_FW_RULE_DIR_IN = 1
Const NET_FW_RULE_DIR_OUT = 2

' Action
Const NET_FW_ACTION_BLOCK = 0
Const NET_FW_ACTION_ALLOW = 1


' Create the FwPolicy2 object.
Dim fwPolicy2
Set fwPolicy2 = CreateObject("HNetCfg.FwPolicy2")

CurrentProfiles = fwPolicy2.CurrentProfileTypes


'// The returned 'CurrentProfiles' bitmask can have more than 1 bit set if multiple profiles 
'//   are active or current at the same time

if ( CurrentProfiles AND NET_FW_PROFILE2_DOMAIN ) then
   WScript.Echo("Domain Firewall Profile is active")
end if

if ( CurrentProfiles AND NET_FW_PROFILE2_PRIVATE ) then
   WScript.Echo("Private Firewall Profile is active")
end if

if ( CurrentProfiles AND NET_FW_PROFILE2_PUBLIC ) then
   WScript.Echo("Public Firewall Profile is active")
end if



' Get the Rules object
Dim RulesObject
Set RulesObject = fwPolicy2.Rules
Dim counterIn, counterOut
counterIn = 0
counterOut= 0
For Each ruleIterate In Rulesobject
    if ruleIterate.Direction = NET_FW_RULE_DIR_IN and ruleIterate.Name = "Custom_Inbound_Rule_KACE" then
		counterIn = 1
	end if
	if ruleIterate.Direction = NET_FW_RULE_DIR_OUT and ruleIterate.Name = "Custom_Outbound_Rule_KACE" then
		counterOut = 1
	end if
Next

if counterIn = 0 then
	Dim NewRuleIn
	Set NewRuleIn = CreateObject("HNetCfg.FWRule")
	NewRuleIn.Name = "Custom_Inbound_Rule_KACE"
	NewRuleIn.Description = "Allow Inbound network traffic over TCP port 52230"
	NewRuleIn.Protocol = NET_FW_IP_PROTOCOL_TCP
	NewRuleIn.LocalPorts = 52230
	NewRuleIn.Direction = NET_FW_RULE_DIR_IN
	NewRuleIn.Enabled = TRUE
	NewRuleIn.Action = NET_FW_ACTION_ALLOW
	RulesObject.Add NewRuleIn
end if

if counterOut = 0 then
	Dim NewRuleOut
	Set NewRuleOut = CreateObject("HNetCfg.FWRule")
	NewRuleOut.Name = "Custom_Outbound_Rule_KACE"
	NewRuleOut.Description = "Allow outbound network traffic over TCP port 52230"
	NewRuleOut.Protocol = NET_FW_IP_PROTOCOL_TCP
	NewRuleOut.LocalPorts = 52230
	NewRuleOut.Direction = NET_FW_RULE_DIR_OUT
	NewRuleOut.Enabled = TRUE
	NewRuleOut.Action = NET_FW_ACTION_ALLOW
	RulesObject.Add NewRuleOut
end if

' Print all the rules in currently active firewall profiles.
WScript.Echo("Rules:")

For Each rule In Rulesobject
    'if rule.Profiles And CurrentProfiles then
			if (rule.Protocol = NET_FW_IP_PROTOCOL_TCP or rule.Protocol = NET_FW_IP_PROTOCOL_UDP) and rule.Enabled = true then
				if rule.Name <> "Custom_Outbound_Rule_KACE" and rule.Name <> "Custom_Inbound_Rule_KACE" then
					WScript.Echo("  Rule Name:          " & rule.Name)
					WScript.Echo("  Direction:    " & rule.Direction)		
					WScript.Echo("  status:       " &  rule.Enabled )		
					rule.Enabled = false
				end if
			end if
	'end if
Next