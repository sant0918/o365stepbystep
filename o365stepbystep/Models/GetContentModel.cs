// Copyright � Microsoft Corporation.  All Rights Reserved.
// This code released under the terms of the 
// Microsoft Public License (MS-PL, http://opensource.org/licenses/ms-pl.html.)

/*
Sample Code is provided for the purpose of illustration only and is not intended 
to be used in a production environment. THIS SAMPLE CODE AND ANY RELATED INFORMATION 
ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS 
FOR A PARTICULAR PURPOSE. We grant You a nonexclusive, royalty-free right to 
use and modify the Sample Code and to reproduce and distribute the object code 
form of the Sample Code, provided that. You agree: (i) to not use Our name, logo, 
or trademarks to market Your software product in which the Sample Code is embedded; 
(ii) to include a valid copyright notice on Your software product in which the 
Sample Code is embedded; and (iii) to indemnify, hold harmless, and defend Us 
and Our suppliers from and against any claims or lawsuits, including attorneys� 
fees, that arise or result from the use or distribution of the Sample Code
*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace o365stepbystep.Models
{

    public class GetContentModel
    {
        public Content[] content { get; set; }
    }

    public class Content 
    {
        public string Id { get; set; }
        public int RecordType { get; set; }
        public DateTime CreationTime { get; set; }
        public string Operation { get; set; }
        public string OrganizationId { get; set; }
        public int UserType { get; set; }
        public string UserKey { get; set; }
        public string Workload { get; set; }
        public string ResultStatus { get; set; }
        public string ObjectId { get; set; }
        public string UserId { get; set; }
        public string ClientIp { get; set; }
        public bool ExternalAccess { get; set; }
        public string OrganizationName { get; set; }
        public string OriginatingServer { get; set; }
        public Parameter[] Parameters { get; set; }
    }
    
    public class Parameter
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }
    
    public class SPContent 
    {
        public DateTime CreationTime { get; set; }
        public string Id { get; set; }
        public string Operation { get; set; }
        public string OrganizationId { get; set; }
        public int RecordType { get; set; }
        public string UserKey { get; set; }
        public int UserType { get; set; }
        public int Version { get; set; }
        public string Workload { get; set; }
        public string ClientIP { get; set; }
        public string ObjectId { get; set; }
        public string UserId { get; set; }
        public string EventSource { get; set; }
        public string ItemType { get; set; }
        public string Site { get; set; }
        public string UserAgent { get; set; }
        public Modifiedproperty[] ModifiedProperties { get; set; }
        public string WebId { get; set; }
        public string EventData { get; set; }
        public string TargetUserOrGroupType { get; set; }
        public string TargetUserOrGroupName { get; set; }

    }

    public class CommonSchema
    {
        public string Id { get; set; }
        public int RecordType { get; set; }
        public DateTime CreationTime { get; set; }
        public string Operation { get; set; }
        public string OrganizationId { get; set; }
        public int UserType { get; set; }
        public string UserKey { get; set; }
        public string Workload { get; set; }
        public string ResultStatus { get; set; }
        public string ObjectId { get; set; }
        public string UserId { get; set; }
        public string ClientIp { get; set; }
    }



    public class Modifiedproperty
    {
        public string Name { get; set; }
        public string NewValue { get; set; }
    }


    public class SharepointFile
    {
        public DateTime CreationTime { get; set; }
        public string Id { get; set; }
        public string Operation { get; set; }
        public string OrganizationId { get; set; }
        public int RecordType { get; set; }
        public string UserKey { get; set; }
        public int UserType { get; set; }
        public int Version { get; set; }
        public string Workload { get; set; }
        public string ClientIP { get; set; }
        public string ObjectId { get; set; }
        public string UserId { get; set; }
        public string EventSource { get; set; }
        public string ItemType { get; set; }
        public string ListItemUniqueId { get; set; }
        public string Site { get; set; }
        public string UserAgent { get; set; }
        public string WebId { get; set; }
        public string SourceFileExtension { get; set; }
        public string SiteUrl { get; set; }
        public string SourceFileName { get; set; }
        public string SourceRelativeUrl { get; set; }
    }


    public class SharepointDlp
    {
        public DateTime CreationTime { get; set; }
        public string Id { get; set; }
        public string Operation { get; set; }
        public string OrganizationId { get; set; }
        public int RecordType { get; set; }
        public string UserKey { get; set; }
        public int UserType { get; set; }
        public int Version { get; set; }
        public string Workload { get; set; }
        public string UserId { get; set; }
        public Policydetail[] PolicyDetails { get; set; }
        public Sharepointmetadata SharePointMetaData { get; set; }
    }

    public class Sharepointmetadata
    {
        public DateTime ItemCreationTime { get; set; }
        public DateTime ItemLastModifiedTime { get; set; }
        public string SiteCollectionGuid { get; set; }
        public string UniqueID { get; set; }
        public string FileName { get; set; }
        public string FileOwner { get; set; }
        public string FilePathUrl { get; set; }
        public string From { get; set; }
        public string SiteCollectionUrl { get; set; }
    }

    public class Policydetail
    {
        public string PolicyId { get; set; }
        public string PolicyName { get; set; }
        public Rule[] Rules { get; set; }
    }

    public class Rule
    {
        public Conditionsmatched ConditionsMatched { get; set; }
        public string RuleId { get; set; }
        public string RuleMode { get; set; }
        public string Severity { get; set; }
    }

    public class Conditionsmatched
    {
        public Sensitiveinformation[] SensitiveInformation { get; set; }
    }

    public class Sensitiveinformation
    {
        public int Confidence { get; set; }
        public int Count { get; set; }
        public string SensitiveType { get; set; }
    }




}
