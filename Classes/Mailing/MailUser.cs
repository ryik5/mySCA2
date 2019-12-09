using System;

namespace ASTA.Classes
{
    class MailUser
    {
        private MailUser _mailUser;
        private string _email;
        private string _name;

        public MailUser()
        {
        }

        public MailUser(string name, string email)
        {
            SetName(name);
            SetEmail(email);
        }

        public MailUser Get()
        {
            return _mailUser;
        }

        public string GetName()
        {
            return _mailUser?._name;
        }
        public string GetEmail()
        {
            return _mailUser?._email;
        }

        public void SetName(string name)
        {
            if (_mailUser != null)
            {
                _mailUser._name = name ?? string.Empty;
            }
            else
            {
                _mailUser = new MailUser
                {
                    _name = name ?? string.Empty
                };
            }
        }
        public void SetEmail(string email)
        {
            if (_mailUser != null)
            {
                _mailUser._email = email;
            }
            else
            {
                _mailUser = new MailUser
                {
                    _email = email
                };
            }
        }

        public override string ToString()
        {
            if (_mailUser != null)
            {
                return $"{_mailUser?._name}\t{_mailUser?._email}";
            }
            else
            { return null; }
        }

        public override bool Equals(object obj)
        {
            if (obj == null || !(obj is MailUser))
                return false;

            MailUser df = obj as MailUser;
            if (df == null)
                return false;

            return ToString().Equals(df.ToString());
        }

        public override int GetHashCode()
        {
            return ToString().GetHashCode();
        }
    }
}
