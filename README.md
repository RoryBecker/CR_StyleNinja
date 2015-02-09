# 'Style Ninja' for CodeRush

Provides StyleCop style hints through the CodeRush CodeIssues interface without the need to install StyleCop. 
These hints/rules will work across languages using user preferences (as stored in CodeRush's style settings) rather that actual StyleCop rules, where applicable.

### Usage 
These rules will work as long as the CodeIssues function is turned on. 

### Options 
This plugin does not provide it's own options screen at the moment. It will however use the preferences set up in CodeRush's "**Editor\Code Style\Identifiers**"  options page.

### Rules 

The following rules are implemented at this time:

 * Naming (!CodeRush settings based).
   * Fields start with [Your Preference Here]
   * Locals start with [Your Preference Here]
   * Params start with [Your Preference Here]
 * SA11XX - Readability
   * SA1100 - Do not prefix calls with Base or MyBase unless a local implementation exists.
   * SA1101 - Prefix local calls this or Me.
 * SA13XX - Naming
   * SA1300 - Namespace, Class, Structure, Enum, Delegate, Event, Method and Property names must begin with an uppercase character.
   * SA1302 - Interface names must begin with I
   * SA1304 - Non-Private Readonly fields must begin with an uppercase character
   * SA1305 - Field names must not use Hungarian Notation.
   * SA1306 - Internal Field names must begin with a lowercase character
   * SA1307 - External Field names must begin with an uppercase character
   * SA1308 - Variable names must not be prefixed [s_] or [m_]
   * SA1309 - Field names should not begin with [_]
   * SA1310 - Field names should not contain [_]
 * SA14XX - Maintainability
   * SA1400 - Access Modifier Must Be Declared.
   * SA1401 - Fields must be Private
   * SA1402 - (Partial) Files must only contain a single Class
   * SA1403 - (Partial) Files must only contain a single Namespace
   * SA1404 - (No fix Possible) SuppressMessage attribute does not include a justification
   * SA1405 - (No fix Possible) Debug.Assert does not include a message
   * SA1406 - (Impossible Rule) Debug.Fail() does not specify a descriptive message
   * SA1409 - Remove Unnecessary Code
 * SA15XX - Layout
   * SA1507 - Multiple Blank Lines are Bad.
= History = 
 * Build 1342 - New rules implemented.
   * Readability rules - SA1101 and SA1100
   * Layout rule - SA1507.
 * Build 940 
   * Code reduced through more Refactoring. Less to maintain FTW! 
 * Build 928  
   * Major Refactoring - Check everything :)
   * Previously Supported rules now have fixes as well.
 * Build 914 - Fixed SA1307 - Previously required an initial lowercase char rather than uppercase.
 * Build 909 - Added SA1300 series of rules 
   * SA1301 is missing because it is unused at this time. 
   * SA1303 is missing because it was unclear how it should work.
 * Build 895 - Initial test build with 5 rules

