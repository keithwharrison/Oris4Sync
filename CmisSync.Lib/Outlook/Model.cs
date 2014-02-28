using Microsoft.Office.Interop.Outlook;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace CmisSync.Lib.Outlook
{
    public class Email
    {
        public string messageID { get; set; }
        public string references { get; set; }
        public DateTime receivedDate { get; set; }
        public DateTime sentDate { get; set; }
        public bool attachmentOnly { get; set; }
        public string subject { get; set; }
        public string body { get; set; }
        public DateTime dateCreated { get; set; }
        public string dataHash { get; set; }
        public string folderPath { get; set; }
        public string inReplyTo { get; set; }
        public List<EmailContact> emailContacts { get; set; }
        public List<EmailAttachment> attachments { get; set; }
        public DateTime lastModified { get; set; }
        public long key { get; set; }
        [JsonIgnore]
        public string entryID { get; set; }
        [JsonIgnore]
        public MailItem mailItem { get; set; }
    }

    public class EmailContact
    {
        public string emailContactType { get; set; }
        public string emailAddress { get; set; }
    }

    public class EmailAttachment
    {
        public string dataHash { get; set; }
        [JsonProperty("emailDatahash")]
        public string emailDataHash { get; set; }
        public string name { get; set; }
        public string fileName { get; set; }
        public long fileSize { get; set; }
        [JsonIgnore]
        public string folderPath { get; set; }
        [JsonIgnore]
        public string tempFilePath { get; set; }
        [JsonIgnore]
        public string entryID { get; set; }
        [JsonIgnore]
        public Attachment attachment { get; set; }
    }
    
    public class OAuth
    {
        public string value { get; set; }
        public object expiration { get; set; }
        public string tokenType { get; set; }
        public List<string> scope { get; set; }
        public OAuthUserInfo additionalInformation { get; set; }
        public long expiresIn { get; set; }
        public bool expired { get; set; }
    }

    public class OAuthUserInfo
    {
        public string username { get; set; }
        public long systemUserKey { get; set; }
        public long individualKey { get; set; }
        public string systemUserRoleName { get; set; }
    }

    public class OutlookFolder
    {
        public string entryId { get; set; }
        public string name { get; set; }
        public string folderPath { get; set; }
        private List<OutlookFolder> _children = new List<OutlookFolder>();
        public List<OutlookFolder> children { get { return _children; } }
    }
}
