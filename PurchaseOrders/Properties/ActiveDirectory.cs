using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.DirectoryServices.ActiveDirectory;
using System.Linq;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;

namespace BSGlobals
{
    public class ActiveDirectory
    {

        public List<BSOUClass> BSOUList;
        public List<ADUserClass> ADUserList;
        public WindowsIdentity CurrentUser;
        public DirectoryEntry DirectoryEntry;
        public string Username;
        public bool UsernameFound;

        public ActiveDirectory()
        {
            ADUserList = new List<ADUserClass>();
            BSOUList = new List<BSOUClass>();
            CurrentUser = WindowsIdentity.GetCurrent();
            UsernameFound = GetADUser(CurrentUser.Name);
        }

        public bool GetADUser(string loginUsername)
        {
            // Obtain the specified Active Directory user.  The user attributes we'll be interested include
            //    First and Last name
            //    User name
            //    Active/Inactive status
            //    All credentials (including both BSOU_ and non-BSOU_ credentials)
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, "buffnews.com")) // Note that we're hard-coding "Buffnews.com" here
                {
                    Username = StripOffUsername(loginUsername);
                    using (var searcher = new PrincipalSearcher(new UserPrincipal(context) { SamAccountName = Username } ))
                    {
                        // Return the first AD entry found for this user.  There should only be a single entry per username!!!!
                        PrincipalSearchResult<Principal> principal = searcher.FindAll();
                        foreach (var result in principal)
                        {
                            DirectoryEntry = result.GetUnderlyingObject() as DirectoryEntry;
                            PopulateADUser(DirectoryEntry);
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                // TBD
                return false;
            }
        }

        private string StripOffUsername(string username)
        {
            // Usernames consist of DOMAIN\Username.  Split at the slash and return only the user's name
            string[] name = username.Split('\\');
            return(name.Last());
        }

        public void GetAllADUsers()
        {
            // Obtain a list of ALL Active Directory users.  The user attributes we'll be interested include
            //    First and Last names
            //    User names
            //    Active/Inactive status
            //    All credentials (including both BSOU_ and non-BSOU_ credentials)
            // WARNING:  THIS CAN TAKE A LONG TIME...
            try
            {
                using (var context = new PrincipalContext(ContextType.Domain, "buffnews.com"))
                {
                    using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                    {
                        PrincipalSearchResult<Principal> principal = searcher.FindAll();
                        foreach (var result in principal)
                        {
                            DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                            PopulateADUser(de);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // TBD
            }
        }

        private void PopulateADUser(DirectoryEntry de)
        {
            // Add this user and his/her credentials to the user and BSOU lists.
            try
            {
                string firstname = "";
                if (de.Properties["givenName"].Value != null)
                {
                    firstname = de.Properties["givenName"].Value.ToString();
                }
                string lastname = "";
                if (de.Properties["sn"].Value != null)
                {
                    lastname = de.Properties["sn"].Value.ToString();
                }
                string username = "";
                if (de.Properties["samAccountName"].Value != null)
                {
                    username = de.Properties["samAccountName"].Value.ToString();
                }
                //0x0002 && user account control flags will determine if user account is enabled or disabled.
                bool isenabled = false;
                if (de.Properties["userAccountControl"].Value != null)
                {
                    int flags = (int)de.Properties["userAccountControl"].Value;
                    isenabled = ((flags & 0x0002) == 0);
                }
                ADUserList.Add(new ADUserClass(firstname, lastname, username, isenabled));

                // Get all attributes associated with this user and add them to their entry
                List<string> memberoflist = GetUserCredentials(de);
                //List<string> memberoflist = GetUserMemberOf(de);

                memberoflist.Sort();
                for (int i = 0; i < memberoflist.Count; i++)
                {
                    string bsoucredential = memberoflist[i];
                    ADUserClass aduser = ADUserList[ADUserList.Count - 1];
                    aduser.AddCredential(bsoucredential);
                    // Add this user to the BSOU list
                    int index = BSOUList.FindIndex(x => x.Credential == bsoucredential);
                    if (index < 0)
                    {
                        bool isbsoucredential = (bsoucredential.Left(5) == "BSOU_");
                        BSOUList.Add(new BSOUClass(bsoucredential, isbsoucredential));
                        index = BSOUList.Count - 1;
                    }
                    BSOUList[index].AddADUser(aduser);
                }
            }
            catch (Exception ex)
            {
                // TBD
            }
        }

        private static List<string> GetAdminGroup(string group)
        {
            List<string> admingroup = new List<string>();
            DirectoryEntry localMachine = new DirectoryEntry($"LDAP://{group.Replace("/", "\\/")}");
            DirectoryEntry admGroup = localMachine.Children.Find("administrators", "group");
            string members = (string)admGroup.Invoke("members", null);

            foreach (object groupMember in members)
            {
                DirectoryEntry member = new DirectoryEntry(groupMember);
                Console.WriteLine(member.Name);
                admingroup.Add(member.Name);
            }
            return (admingroup);
        }

        private static List<string> GetUserCredentials(DirectoryEntry de)
        {
            var groups = new List<string>();
            try
            {
                // Retrieve only the memberOf attributes from the user.
                //  This includes ALL attributes, not just BSOU_ attributes
                de.RefreshCache(new[] { "memberOf" });

                var credentials = de.Properties["memberOf"];
                foreach (string group in credentials)
                {
                    var groupDe = new DirectoryEntry($"LDAP://{group.Replace("/", "\\/")}");
                    groupDe.RefreshCache(new[] { "cn" });
                    string credential = groupDe.Properties["cn"].Value as string;
                    string distinguishedname = groupDe.Properties["distinguishedName"].Value as string;

                    groups.Add(credential);
                }
            }
            catch (Exception ex)
            {
                // TBD
            }
            return groups;
        }

        // TBD - We may want to recursively get users within groups!!!!
        private static List<string> GetUserMemberOf(DirectoryEntry de)
        {
            var groups = new List<string>();
            try
            {
                // Retrieve only the memberOf attributes from the user
                de.RefreshCache(new[] { "memberOf" });

                var memberOf = de.Properties["memberOf"];
                foreach (string group in memberOf)
                {
                    var groupDe = new DirectoryEntry($"LDAP://{group.Replace("/", "\\/")}");
                    groupDe.RefreshCache(new[] { "cn" });
                    groups.Add(groupDe.Properties["cn"].Value as string);
                }

            }
            catch (Exception ex)
            {
                // TBD
            }
            return groups;
        }
    }

    public class BSOUClass
    {
        private string _Credential;
        private bool _ISBSOUCredential;
        private List<ADUserClass> _ADUser;

        public BSOUClass(string credential, bool isbsoucredential)
        {
            try
            {
                _Credential = credential;
                _ISBSOUCredential = isbsoucredential;
                _ADUser = new List<ADUserClass>();
            }
            catch (Exception ex)
            {
                // TBD
            }
        }

        public void AddADUser(ADUserClass aduser)
        {
            try
            {
                _ADUser.Add(aduser);
            }
            catch (Exception ex)
            {
                // TBD
            }
        }

        public string Credential { get { return this._Credential; } }
        public bool IsBSOUCredential { get { return this._ISBSOUCredential; } }
        public List<ADUserClass> ADUser { get { return this._ADUser; } }

    }

    public class ADUserClass
    {
        private string _FirstName;
        private string _LastName;
        private string _UserName;
        private bool _AccountEnabled;
        private List<string> _Credentials;

        public ADUserClass(string firstname, string lastname, string username, bool accountenabled)
        {
            _FirstName = firstname;
            _LastName = lastname;
            _UserName = username;
            _AccountEnabled = accountenabled;
            _Credentials = new List<string>();
        }

        public ADUserClass()
        {
            _FirstName = "<undefined>";
            _LastName = "<undefined>";
            _UserName = "<undefined>";
            _AccountEnabled = false;
            _Credentials = new List<string>();
        }

        public void PopulateUser(string firstname, string lastname, string username, bool accountenabled)
        {
            try
            {
                _FirstName = firstname;
                _LastName = lastname;
                _UserName = username;
                _AccountEnabled = accountenabled;
                _Credentials = new List<string>();
            }
            catch (Exception ex)
            {
                // TBD
            }
        }

        public void AddCredential(string credential)
        {
            try
            {
                _Credentials.Add(credential);
            }
            catch (Exception ex)
            {
                // TBD
            }
        }

        public bool HasBSOUCredentials()
        {
            // Return true if this user has any BSOU credentials
            try
            {
                for (int i = 0; i < _Credentials.Count; i++)
                {
                    if (IsBSOUCredential(_Credentials[i]))
                    {
                        return (true);
                    }
                }
            }
            catch (Exception ex)
            {
                // TBD
            }
            return (false);
        }

        public bool IsBSOUCredential(string credential)
        {
            try
            {
                return (credential.Left(5) == "BSOU_");
            }
            catch (Exception ex)
            {
                // TBD
            }
            return (false);
        }

        public string FirstName { get { return this._FirstName; } }
        public string LastName { get { return this._LastName; } }
        public string UserName { get { return this._UserName; } }
        public bool AccountEnabled { get { return this._AccountEnabled; } }
        public List<string> Credentials { get { return this._Credentials; } }
    }

}
