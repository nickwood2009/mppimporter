using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;

namespace ADC.MppImport
{
    public static class SDKHelpers
    {

        #region Data Extraction
        public static Guid ExtractRecordIDFromRecordURL(string RecordURL)
        {
            RecordURL = RecordURL.Substring(RecordURL.IndexOf("?") + 1);
            string[] strsplit = RecordURL.Split(new string[] { "&" }, StringSplitOptions.RemoveEmptyEntries);
            string idstr = strsplit.Where(x => x.ToLower().Contains("id=")).FirstOrDefault();
            idstr = idstr.Replace("id=", "");
            idstr = idstr.Replace("amp;", "");
            Guid ret = Guid.Empty;
            Guid.TryParse(idstr, out ret);
            return ret;
        }
        public static int ExtractEntityTypeCodeFromRecordURL(string RecordURL)
        {
            RecordURL = RecordURL.Substring(RecordURL.IndexOf("?") + 1);
            string[] strsplit = RecordURL.Split(new string[] { "&" }, StringSplitOptions.RemoveEmptyEntries);
            string idstr = strsplit.Where(x => x.ToLower().Contains("etc=")).FirstOrDefault();
            idstr = idstr.Replace("etc=", "");
            idstr = idstr.Replace("amp;", "");
            int ret = 0;
            Int32.TryParse(idstr, out ret);
            return ret;
        }
        public static string ExtractCRMBaseUrl(this OrganizationServiceProxy svc)
        {
            string ret = "";
            string FullSvcUrl = svc.ServiceManagement.CurrentServiceEndpoint.Address.Uri.ToString();
            ret = FullSvcUrl.Replace("/XRMServices/2011/Organization.svc", "");
            return ret;
        }
        #endregion

        #region General Retrieve Helpers
        public static void RetrieveOptionSetValues(IOrganizationService service, string entityName, string OptionsetAttributeName, out Dictionary<string, int> optionsetvalues)
        {
            optionsetvalues = null;
            RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = OptionsetAttributeName,
                RetrieveAsIfPublished = true,
            };

            // Execute the request.
            RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
            Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata retrievedPicklistAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
            // Get the current options list for the retrieved attribute.
            OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
            optionsetvalues = new Dictionary<string, int>();
            foreach (OptionMetadata oMD in optionList)
            {
                optionsetvalues.Add(oMD.Label.UserLocalizedLabel.Label, (int)oMD.Value);
            }
        }
        public static void RetrieveGlobalOptionSetValues(IOrganizationService service, string OptionsetName, out Dictionary<string, int> optionsetvalues)
        {
            optionsetvalues = null;
            RetrieveOptionSetRequest retrieveOptionSetRequest = new RetrieveOptionSetRequest { Name = OptionsetName };

            // Execute the request.
            RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);
            OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
            OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();

            optionsetvalues = new Dictionary<string, int>();
            foreach (OptionMetadata oMD in optionList)
            {
                optionsetvalues.Add(oMD.Label.UserLocalizedLabel.Label, (int)oMD.Value);
            }
        }
        public static Dictionary<string, int> RetrieveOptionSetValues(IOrganizationService service, string entityName, string OptionsetName)
        {
            Dictionary<string, int> optionsetvalues = new Dictionary<string, int>();

            try
            {
                RetrieveAttributeRequest retrieveAttributeRequest = new RetrieveAttributeRequest
                {
                    EntityLogicalName = entityName,
                    LogicalName = OptionsetName,
                    RetrieveAsIfPublished = true,
                };

                // Execute the request.
                RetrieveAttributeResponse retrieveAttributeResponse = (RetrieveAttributeResponse)service.Execute(retrieveAttributeRequest);
                Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata retrievedPicklistAttributeMetadata = (Microsoft.Xrm.Sdk.Metadata.PicklistAttributeMetadata)retrieveAttributeResponse.AttributeMetadata;
                // Get the current options list for the retrieved attribute.
                OptionMetadata[] optionList = retrievedPicklistAttributeMetadata.OptionSet.Options.ToArray();
                optionsetvalues = new Dictionary<string, int>();
                foreach (OptionMetadata oMD in optionList)
                {
                    optionsetvalues.Add(oMD.Label.UserLocalizedLabel.Label, (int)oMD.Value);
                }
            }
            catch (Exception ex) { }

            return optionsetvalues;
        }
        public static Dictionary<string, int> RetrieveGlobalOptionSetValues(IOrganizationService service, string OptionsetName)
        {
            Dictionary<string, int> optionsetvalues = new Dictionary<string, int>();
            RetrieveOptionSetRequest retrieveOptionSetRequest = new RetrieveOptionSetRequest { Name = OptionsetName };

            // Execute the request.
            RetrieveOptionSetResponse retrieveOptionSetResponse = (RetrieveOptionSetResponse)service.Execute(retrieveOptionSetRequest);
            OptionSetMetadata retrievedOptionSetMetadata = (OptionSetMetadata)retrieveOptionSetResponse.OptionSetMetadata;
            OptionMetadata[] optionList = retrievedOptionSetMetadata.Options.ToArray();

            optionsetvalues = new Dictionary<string, int>();
            foreach (OptionMetadata oMD in optionList)
            {
                optionsetvalues.Add(oMD.Label.UserLocalizedLabel.Label, (int)oMD.Value);
            }

            return optionsetvalues;
        }

        public static Entity GenFetchXML(this IOrganizationService service, string FetchXML)
        {
            EntityCollection result = service.RetrieveMultiple(new FetchExpression(FetchXML));
            Entity firstEntity = result.Entities.FirstOrDefault();
            return firstEntity;
        }
        public static List<Entity> GenFetchXMLList(this IOrganizationService service, string FetchXML)
        {
            EntityCollection result = service.RetrieveMultiple(new FetchExpression(FetchXML));
            return result.Entities.ToList();
        }
        public static string OptionsetValueToLabel(IOrganizationService service, string entityName, string OptionsetName, int TargetValue)
        {
            string ret = "";
            try
            {
                ret = RetrieveOptionSetValues(service, entityName, OptionsetName).Where(x => x.Value == TargetValue).Select(x => x.Key).FirstOrDefault();
            }
            catch (Exception ex) { }
            return ret;
        }

        public static string CreateXml(string xml, string cookie, int page, int count)
        {
            StringReader stringReader = new StringReader(xml);
            XmlTextReader reader = new XmlTextReader(stringReader);

            // Load document
            XmlDocument doc = new XmlDocument();
            doc.Load(reader);

            return CreateXml(doc, cookie, page, count);
        }

        public static string CreateXml(XmlDocument doc, string cookie, int page, int count)
        {
            XmlAttributeCollection attrs = doc.DocumentElement.Attributes;

            if (cookie != null)
            {
                XmlAttribute pagingAttr = doc.CreateAttribute("paging-cookie");
                pagingAttr.Value = cookie;
                attrs.Append(pagingAttr);
            }

            XmlAttribute pageAttr = doc.CreateAttribute("page");
            pageAttr.Value = System.Convert.ToString(page);
            attrs.Append(pageAttr);

            XmlAttribute countAttr = doc.CreateAttribute("count");
            countAttr.Value = System.Convert.ToString(count);
            attrs.Append(countAttr);

            StringBuilder sb = new StringBuilder(1024);
            StringWriter stringWriter = new StringWriter(sb);

            XmlTextWriter writer = new XmlTextWriter(stringWriter);
            doc.WriteTo(writer);
            writer.Close();

            return sb.ToString();
        }

        public static List<Entity> PageCRMFetchXML(string fetchXML, IOrganizationService _orgService, int numPerPage, bool Verbose)
        {
            List<Entity> ret = new List<Entity>();
            int pageNumber = 1;
            int recordCount = 0;
            string pagingCookie = null;
            while (true)
            {
                // Build fetchXml string with the placeholders.
                string xml = CreateXml(fetchXML, pagingCookie, pageNumber, numPerPage);

                // Excute the fetch query and get the xml result.
                RetrieveMultipleRequest fetchRequest1 = new RetrieveMultipleRequest
                {
                    Query = new FetchExpression(xml)
                };
                EntityCollection returnCollection = ((RetrieveMultipleResponse)_orgService.Execute(fetchRequest1)).EntityCollection;
                ret.AddRange(returnCollection.Entities);
                if (Verbose) Console.WriteLine("CRM Records Returned: " + returnCollection.Entities.Count.ToString() + " Total: " + ret.Count.ToString());
                if (returnCollection.MoreRecords)
                {

                    // Increment the page number to retrieve the next page.
                    pageNumber++;

                    // Set the paging cookie to the paging cookie returned from current results.                            
                    pagingCookie = returnCollection.PagingCookie;
                }
                else
                {
                    // If no more records in the result nodes, exit the loop.
                    break;
                }

            }

            return ret;
        }
        #endregion

        #region Attribute Extensions
        public static bool AttrKeyFound(this Entity entity, string key)
        {
            bool ret = false;
            try
            {
                if (entity.Attributes.ContainsKey(key))
                    ret = true;
            }
            catch (Exception ex)
            {
            }
            return ret;
        }
        public static Guid EntRefToIdGuid(this object obj)
        {
            Guid ret = Guid.Empty;
            try
            {
                EntityReference EF = obj as EntityReference;
                if (EF != null)
                {
                    ret = EF.Id;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntRefToIdGuid:{0}", ex.ToString()));
            }
            return ret;
        }
        public static Guid AttributeToIdGuid(this object obj)
        {
            Guid ret = Guid.Empty;
            if (obj.GetType() == typeof(Guid))
            {
                ret = (Guid)obj;
            }
            else if (obj.GetType() == typeof(EntityReference))
            {
                ret = (obj as EntityReference).Id;
            }
            else if (obj.GetType() == typeof(AliasedValue))
            {
                AliasedValue AVal = obj as AliasedValue;
                if (AVal != null)
                {
                    if (AVal.Value.GetType() == typeof(EntityReference))
                        ret = (AVal.Value as EntityReference).Id;
                    else if (AVal.Value.GetType() == typeof(Guid))
                        ret = (Guid)AVal.Value;
                }
            }
            return ret;
        }
        public static DateTime? AttributeToDateTimeNullable(this object obj)
        {
            DateTime? ret = null;
            if (obj.GetType() == typeof(AliasedValue))
            {
                AliasedValue AVal = obj as AliasedValue;
                if (AVal != null)
                {
                    ret = new DateTime(((DateTime)AVal.Value).Ticks, DateTimeKind.Utc);
                }
            }
            else
            {
                ret = new DateTime(((DateTime)obj).Ticks, DateTimeKind.Utc);
            }
            return ret;
        }
        public static List<Guid> AttributeToActivityPartyIdList(this object obj)
        {
            List<Guid> ret = new List<Guid>();
            if (obj.GetType() == typeof(EntityCollection))
            {
                foreach (Entity E in ((EntityCollection)obj).Entities)
                {
                    ret.Add(E.EntityAttributeToGuid("partyid"));
                }
            }
            return ret;
        }
        public static List<Entity> AttributeToActivityPartyEntityList(this object obj)
        {
            List<Entity> ret = new List<Entity>();
            if (obj.GetType() == typeof(EntityCollection))
            {
                foreach (Entity E in ((EntityCollection)obj).Entities)
                {
                    ret.Add(E);
                }
            }
            return ret;
        }
        public static string AttributeToString(this object obj)
        {
            string ret = string.Empty;
            if (obj.GetType() == typeof(AliasedValue))
            {
                AliasedValue AVal = obj as AliasedValue;
                if (AVal != null)
                {
                    if (AVal.Value.GetType() == typeof(EntityReference))
                        ret = (AVal.Value as EntityReference).Name;
                    else
                        ret = AVal.Value.ToString();
                }
            }
            else
            {
                ret = obj.ToString();
            }
            return ret;
        }
        public static int AttributeToInt(this object obj)
        {
            int ret = 0;
            if (obj.GetType() == typeof(AliasedValue))
            {
                AliasedValue AVal = obj as AliasedValue;
                if (AVal != null)
                {
                    ret = Convert.ToInt16(AVal.Value);
                }
            }
            else
            {
                ret = Convert.ToInt16(obj);
            }
            return ret;
        }
        public static OptionSetValue AttributeOptionSetValue(this object obj)
        {
            OptionSetValue ret = null;
            try
            {
                if (obj.GetType() == typeof(AliasedValue))
                {
                    AliasedValue AV = obj as AliasedValue;
                    if (AV != null)
                    {
                        OptionSetValue OSVal = AV.Value as OptionSetValue;
                        if (OSVal != null) ret = OSVal;
                    }
                }
                else if (obj.GetType() == typeof(OptionSetValue))
                {
                    OptionSetValue OSVal = obj as OptionSetValue;
                    if (OSVal != null) ret = OSVal;
                }
            }
            catch (Exception ex)
            {
                //ex.LogError();
            }
            return ret;
        }
        public static bool AttributeToBool(this object obj)
        {
            bool ret = false;
            try
            {
                if (obj.GetType() == typeof(AliasedValue))
                {
                    AliasedValue AV = obj as AliasedValue;
                    if (AV != null)
                    {
                        ret = Convert.ToBoolean(AV.Value);
                    }
                }
                else if (obj.GetType() == typeof(bool))
                {
                    ret = Convert.ToBoolean(obj);
                }
            }
            catch (Exception ex)
            {
            }
            return ret;
        }
        public static decimal AttributeToDecimal(this object obj)
        {
            decimal ret = 0;
            if (obj.GetType() == typeof(AliasedValue)) obj = AttributeAliasedValueToValue(obj);
            if (obj.GetType() == typeof(Money)) ret = (obj as Money).Value;
            else ret = Convert.ToDecimal(obj);

            return ret;
        }
        public static double AttributeToDouble(this object obj)
        {
            double ret = 0;
            if (obj.GetType() == typeof(AliasedValue))
            {
                AliasedValue AVal = obj as AliasedValue;
                if (AVal != null)
                {
                    ret = Convert.ToDouble(AVal.Value);
                }
            }
            else
            {
                ret = Convert.ToDouble(obj);
            }
            return ret;
        }
        public static DateTime AttributeToDateTime(this object obj)
        {
            DateTime ret = DateTime.Now;
            if (obj.GetType() == typeof(AliasedValue))
            {
                AliasedValue AVal = obj as AliasedValue;
                if (AVal != null)
                {
                    ret = Convert.ToDateTime(AVal.Value);
                }
            }
            else
            {
                ret = Convert.ToDateTime(obj);
            }
            return ret;
        }
        public static object GetAttributeVal(this Entity obj, string attrName)
        {
            object ret = null;
            if (obj.Attributes.Contains(attrName))
            {
                ret = AttributeAliasedValueToValue(obj.Attributes[attrName]);
            }
            return ret;
        }
        public static string GetAttrToFormattedString(IOrganizationService service, Entity Ent, string fieldName, string EntityName)
        {
            string ret = "";
            try
            {
                object targetValue = Ent.GetAttributeVal(fieldName);
                if (targetValue.GetType() == typeof(string)) ret = targetValue.ToString();
                else if (targetValue.GetType() == typeof(Money))
                {
                    Money val = targetValue as Money;
                    ret = val.Value.ToFormattedMoneyString();
                }
                else if (targetValue.GetType() == typeof(bool))
                {
                    ret = Convert.ToBoolean(targetValue).ToString();
                }
                else if (targetValue.GetType() == typeof(decimal))
                {
                    ret = Convert.ToDecimal(targetValue).ToString();
                }
                else if (targetValue.GetType() == typeof(double))
                {
                    ret = Convert.ToDouble(targetValue).ToString();
                }
                else if (targetValue.GetType() == typeof(int))
                {
                    ret = Convert.ToInt32(targetValue).ToString();
                }
                else if (targetValue.GetType() == typeof(Guid))
                {
                    ret = targetValue.ToString();
                }
                else if (targetValue.GetType() == typeof(EntityReference))
                {
                    EntityReference val = targetValue as EntityReference;
                    ret = val.Id.ToString();
                }
                else if (targetValue.GetType() == typeof(DateTime))
                {
                    DateTime val = (DateTime)targetValue;
                    ret = val.ToFormattedDateTimeString();
                }
                else if (targetValue.GetType() == typeof(OptionSetValue))
                {
                    OptionSetValue OSVal = targetValue as OptionSetValue;
                    if (fieldName.Contains("."))
                    {
                        string[] strSplit = fieldName.Split(new char[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
                        EntityName = strSplit[0];
                        fieldName = strSplit[1];
                    }
                    if (OSVal != null) ret = OptionsetValueToLabel(service, EntityName, fieldName, OSVal.Value);
                }
            }
            catch (Exception ex)
            {
            }
            return ret;
        }
        public static object AttributeAliasedValueToValue(object val)
        {
            object targetValue = null;
            if (val.GetType() == typeof(AliasedValue))
            {
                AliasedValue AV = val as AliasedValue;
                targetValue = AV.Value;
            }
            else
                targetValue = val;

            return targetValue;
        }
        public static string AttributeKVPairToString(this KeyValuePair<string, object> attrItem)
        {
            string ret = "";
            object targetValue = null;
            if (attrItem.Value != null)
            {
                targetValue = AttributeAliasedValueToValue(attrItem.Value);

                if (targetValue.GetType() == typeof(string)) ret = targetValue.ToString();
                else if (targetValue.GetType() == typeof(Money))
                {
                    Money val = targetValue as Money;
                    ret = val.Value.ToFormattedMoneyString();
                }
                else if (targetValue.GetType() == typeof(bool))
                {
                    ret = Convert.ToBoolean(targetValue).ToString();
                }
                else if (targetValue.GetType() == typeof(decimal))
                {
                    ret = Convert.ToDecimal(targetValue).ToString();
                }
                else if (targetValue.GetType() == typeof(double))
                {
                    ret = Convert.ToDouble(targetValue).ToString();
                }
                else if (targetValue.GetType() == typeof(int))
                {
                    ret = Convert.ToInt32(targetValue).ToString();
                }
                else if (targetValue.GetType() == typeof(Guid))
                {
                    ret = targetValue.ToString();
                }
                else if (targetValue.GetType() == typeof(EntityReference))
                {
                    EntityReference val = targetValue as EntityReference;
                    ret = val.Id.ToString();
                }
                else if (targetValue.GetType() == typeof(DateTime))
                {
                    DateTime val = (DateTime)targetValue;
                    ret = val.ToFormattedDateTimeString();
                }
                else if (targetValue.GetType() == typeof(OptionSetValue))
                {
                    OptionSetValue OSVal = targetValue as OptionSetValue;
                    if (OSVal != null) ret = OSVal.Value.ToString();
                }
            }

            return ret;
        }
        public static EntityReference AttributeToEntityReference(this object obj)
        {
            EntityReference ret = null;
            if (obj.GetType() == typeof(AliasedValue)) obj = AttributeAliasedValueToValue(obj);
            if (obj.GetType() == typeof(EntityReference)) ret = (obj as EntityReference);
            return ret;
        }
        #endregion

        #region EntityAttribute Extensions
        public static decimal EntityAttributeToMoney(this Entity obj, string attrName)
        {
            decimal ret = 0;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    if (obj.Attributes[attrName].GetType() == typeof(Money))
                    {
                        ret = (obj.Attributes[attrName] as Money).Value;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToMoney attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static void UpsertEntityAttribute(this Entity obj, string attrName, object NewValue)
        {
            if (obj.Attributes.Contains(attrName))
            {
                obj.Attributes[attrName] = NewValue;
            }
            else
            {
                obj.Attributes.Add(new KeyValuePair<string, object>(attrName, NewValue));
            }
        }
        public static string EntityAttributeToString(this Entity obj, string attrName)
        {
            string ret = string.Empty;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName)) ret = obj.Attributes[attrName].AttributeToString();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToString attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            if (!string.IsNullOrEmpty(ret)) ret = ret.Trim();
            return ret;
        }
        public static string EntityAttributeToString(this Entity obj, string attrName, string defaultVal)
        {
            string ret = defaultVal;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    string val = obj.Attributes[attrName].AttributeToString();
                    if (!string.IsNullOrEmpty(val)) ret = val;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToString attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static string EntityAttributeOptionSetNameSelected(this Entity obj, string attrName, Dictionary<string, int> fullOptionsetList)
        {
            string ret = string.Empty;
            try
            {
                OptionSetValue currentAttrOSVal = null;
                if (obj != null && obj.Attributes.Contains(attrName)) currentAttrOSVal = obj.Attributes[attrName].AttributeOptionSetValue();
                if (currentAttrOSVal != null) ret = fullOptionsetList.Where(x => x.Value == currentAttrOSVal.Value).Select(x => x.Key).FirstOrDefault();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeOptionSetNameSelected attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static OptionSetValue EntityAttributeOptionSetValue(this Entity obj, string attrName)
        {
            OptionSetValue ret = null;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName)) ret = obj.Attributes[attrName].AttributeOptionSetValue();
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeOptionSetValue attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }

            return ret;
        }
        public static int EntityAttributeOptionSetValueInt(this Entity obj, string attrName)
        {
            int ret = 0;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName)) ret = obj.Attributes[attrName].AttributeOptionSetValue().Value;
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeOptionSetValue attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }

            return ret;
        }
        public static decimal EntityAttributeMoneyToDecimal(this Entity obj, string attrName)
        {
            decimal ret = 0;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    if (obj.Attributes[attrName].GetType() == typeof(AliasedValue))
                    {
                        AliasedValue AV = obj.Attributes[attrName] as AliasedValue;
                        if (AV != null)
                        {
                            if (AV.Value.GetType() == typeof(Money))
                            {
                                ret = (AV.Value as Money).Value;
                            }
                        }
                    }
                    else if (obj.Attributes[attrName].GetType() == typeof(Money))
                    {
                        ret = (obj.Attributes[attrName] as Money).Value;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToMoney attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static decimal EntityAttributeMoneyToDecimal(this Entity obj, string attrName, int RoundedDecPlaces)
        {
            decimal ret = EntityAttributeMoneyToDecimal(obj, attrName);
            ret = Math.Round(ret, RoundedDecPlaces);
            return ret;
        }
        public static int EntityAttributeToInt(this Entity obj, string attrName)
        {
            int ret = 0;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    if (obj.Attributes[attrName].GetType() == typeof(AliasedValue))
                    {
                        AliasedValue AV = obj.Attributes[attrName] as AliasedValue;
                        if (AV != null)
                        {
                            if (AV.Value.GetType() == typeof(int))
                            {
                                ret = Convert.ToInt32(AV.Value);
                            }
                        }
                    }
                    else if (obj.Attributes[attrName].GetType() == typeof(int))
                    {
                        ret = Convert.ToInt32(obj.Attributes[attrName]);
                    }
                }
            }
            catch (Exception ex)
            {
            }

            return ret;
        }
        public static int EntityAttributeAggregateToInt(this Entity obj, string attrName)
        {
            int ret = 0;
            try
            {
                object val = obj.Attributes[attrName];
                if (val == null) return 0;

                do
                {
                    if (val.GetType() == typeof(AliasedValue))
                    {
                        val = (val as AliasedValue).Value;
                    }
                } while (val.GetType() == typeof(AliasedValue));

                if (val.GetType() == typeof(int))
                {
                    ret = Convert.ToInt32(val);
                }
            }
            catch (Exception ex)
            {
            }

            return ret;
        }
        public static decimal EntityAttributeToDecimal(this Entity obj, string attrName)
        {
            decimal ret = 0;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                    ret = obj.Attributes[attrName].AttributeToDecimal();
            }
            catch (Exception ex)
            {
            }

            return ret;
        }
        public static decimal EntityAttributeToDecimal(this Entity obj, string attrName, int RoundedDecPlaces)
        {
            decimal ret = EntityAttributeToDecimal(obj, attrName);
            ret = Math.Round(ret, RoundedDecPlaces);
            return ret;
        }
        public static double EntityAttributeToDouble(this Entity obj, string attrName)
        {
            double ret = 0;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                    ret = obj.Attributes[attrName].AttributeToDouble();
            }
            catch (Exception ex)
            {
            }

            return ret;
        }
        public static bool EntityAttributeToBool(this Entity obj, string attrName)
        {
            bool ret = false;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName)) ret = obj.Attributes[attrName].AttributeToBool();
            }
            catch (Exception ex)
            {
            }

            return ret;
        }
        public static decimal EntityAttributeToMoneyExclGSTPerc(this Entity obj, string attrName, decimal gstPerc)
        {
            decimal ret = 0;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    if (obj.Attributes[attrName].GetType() == typeof(Money))
                    {
                        if (gstPerc > 0 && (obj.Attributes[attrName] as Money).Value > 0)
                        {
                            ret = (obj.Attributes[attrName] as Money).Value / (1 + gstPerc);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToMoneyExclGSTPerc attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static EntityReference EntityAttributeToEntityReference(this Entity obj, string attrName, bool throwMissing = false)
        {
            EntityReference ret = null;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    ret = obj.Attributes[attrName].AttributeToEntityReference();
                }
                else
                {
                    if (throwMissing) throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToEntityReference does not contain attrName: {0}", attrName));
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToEntityReference attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static Guid EntityAttributeToGuid(this Entity obj, string attrName)
        {
            Guid ret = Guid.Empty;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    ret = obj.Attributes[attrName].AttributeToIdGuid();
                }
                else
                {
                    throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToGuid does not contain attrName: {0}", attrName));
                }
            }
            catch (Exception ex)
            {
            }
            return ret;
        }
        public static Guid? EntityAttributeToGuidNull(this Entity obj, string attrName)
        {
            Guid? ret = null;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    ret = obj.Attributes[attrName].AttributeToIdGuid();
                }
                else
                {
                    throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToGuid does not contain attrName: {0}", attrName));
                }
            }
            catch (Exception ex)
            {
            }
            return ret;
        }
        public static List<Guid> EntityAttributeToActivityPartyIdList(this Entity obj, string attrName)
        {
            List<Guid> ret = new List<Guid>();
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    ret = obj.Attributes[attrName].AttributeToActivityPartyIdList();
                }
            }
            catch (Exception ex)
            {
            }
            return ret;
        }
        public static List<Entity> EntityAttributeToActivityPartyEntityList(this Entity obj, string attrName)
        {
            List<Entity> ret = new List<Entity>();
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    ret = obj.Attributes[attrName].AttributeToActivityPartyEntityList();
                }
            }
            catch (Exception ex)
            {
            }
            return ret;
        }

        public static DateTime? EntityAttributeToDateUTC(this Entity obj, string attrName)
        {
            DateTime? ret = null;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    DateTime? dateUTC = obj.Attributes[attrName].AttributeToDateTimeNullable();
                    if (dateUTC != null) ret = dateUTC;
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToString attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static DateTime? EntityAttributeToDate(this Entity obj, string attrName)
        {
            DateTime? ret = null;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    DateTime? dateUTC = obj.Attributes[attrName].AttributeToDateTimeNullable();
                    if (dateUTC != null) ret = dateUTC.ToLocalDateTime();
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToString attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static DateTime? EntityAttributeToDateFailOver(this Entity obj, List<string> dateFields)
        {
            //Iterate each date in order and return the first that has a value
            DateTime? ret = null;
            foreach (string attrFieldName in dateFields)
            {
                ret = EntityAttributeToDate(obj, attrFieldName);
                if (ret.HasValue) break;
            }
            return ret;
        }
        public static DateTime EntityAttributeToDateNonNull(this Entity obj, string attrName)
        {
            DateTime ret = DateTime.MinValue;
            try
            {
                if (obj != null && obj.Attributes.Contains(attrName))
                {
                    DateTime? dateUTC = obj.Attributes[attrName].AttributeToDateTimeNullable();
                    if (dateUTC != null) ret = dateUTC.ToLocalDateTime();
                    else
                    {
                        ret = DateTime.MinValue;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("Exception EntityAttributeToString attrName: {0} - Errpr: {1}", attrName, ex.ToString()));
            }
            return ret;
        }
        public static KeyValuePair<Guid, string> EntityAttributesToLookupKVPair(this Entity obj, string attrKeyName, string attrValName)
        {
            KeyValuePair<Guid, string> ret = new KeyValuePair<Guid, string>(obj.EntityAttributeToGuid(attrKeyName), obj.EntityAttributeToString(attrValName));
            return ret;
        }
        #endregion

        #region Matching
        public static OptionSetValue MatchTargetOptionNameToOSListVal(this Dictionary<string, int> creditAppMaritalOptions, string contMaritalOSName, int TargetDefaultOSVal)
        {
            OptionSetValue ret = new OptionSetValue(TargetDefaultOSVal);

            //Try to match based on string name to lower
            var firstMatchingOS = creditAppMaritalOptions.Where(x => x.Key.Replace(" ", "").ToLower() == contMaritalOSName.Replace(" ", "").ToLower());
            if (firstMatchingOS != null) ret = new OptionSetValue(firstMatchingOS.Select(x => x.Value).FirstOrDefault());

            return ret;
        }
        public static OptionSetValue MatchTargetOSValOSListVal(this Dictionary<string, int> targetOptionSetList, Dictionary<string, int> sourceOptionSetList, int sourceOSVal, int TargetDefaultOSVal)
        {
            OptionSetValue ret = null;
            if (TargetDefaultOSVal > 0) ret = new OptionSetValue(TargetDefaultOSVal);

            //Try to match based on string name to lower
            string sourceOSName = sourceOptionSetList.Where(x => x.Value == sourceOSVal).FirstOrDefault().Key;
            if (!string.IsNullOrEmpty(sourceOSName))
            {
                var firstMatchingOS = targetOptionSetList.Where(x => x.Key.Replace(" ", "").ToLower() == sourceOSName.Replace(" ", "").ToLower());
                if (firstMatchingOS != null) ret = new OptionSetValue(firstMatchingOS.Select(x => x.Value).FirstOrDefault());
            }

            return ret;
        }
        public static bool EntityStringMatch(this Entity PrimEntity, Entity OtherEntity, string PrimField, string OtherField)
        {
            bool ret = false;
            if (PrimEntity.EntityAttributeToString(PrimField).Trim().ToLower() == OtherEntity.EntityAttributeToString(OtherField).Trim().ToLower()) ret = true;
            return ret;
        }
        #endregion

        #region DateTime Extensions
        public static DateTime ToLocalDateTime(this DateTime? val)
        {
            if (val.HasValue)
            {
                return val.ToLocalDateTime();
            }
            else
            {
                return DateTime.MinValue;
            }
        }
        public static DateTime ToUTCTime(this DateTime val)
        {
            DateTime convertedDT = TimeZoneInfo.ConvertTimeToUtc(val);
            return convertedDT;
        }
        public static DateTime? ToUTCDateTime(this DateTime? val)
        {
            if (val.HasValue)
            {
                DateTime convertedDT = TimeZoneInfo.ConvertTimeToUtc(val.Value);
                return convertedDT;
            }
            else
            {
                return null;
            }
        }
        public static DateTime ToHourMinOfDay(this DateTime val, int Hour, int min)
        {
            return new DateTime(val.Year, val.Month, val.Day, Hour, min, 0);
        }
        #endregion

        #region General Extentions
        public static decimal AddGSTPerc(this decimal val, decimal gstPerc)
        {
            return val * (1 + gstPerc);
        }

        /// <summary>
        /// Get the GST amount that would be charged for the current EX GST value
        /// </summary>
        /// <param name="val"></param>
        /// <param name="gstPerc"></param>
        /// <returns></returns>
        public static decimal GSTOnTop(this decimal val, decimal gstPerc)
        {
            if (val == 0) return 0;
            else
                return val * gstPerc;
        }
        public static decimal ToExGST(this decimal val, decimal gstPerc)
        {
            if (val == 0) return 0;
            else
                return val / (1 + gstPerc);
        }

        /// <summary>
        /// Get the GST portion of the current Inc GST value
        /// </summary>
        /// <param name="val"></param>
        /// <param name="gstPerc"></param>
        /// <returns></returns>
        public static decimal ToGSTAmount(this decimal val, decimal gstPerc)
        {
            if (val == 0) return 0;
            else
                return val / ((1 + gstPerc) * 10);
        }
        public static decimal ToNegativeValue(this decimal val)
        {
            if (val < 0) return val;
            else
            {
                return val * -1;
            }
        }
        public static decimal ToNegPosOppositeValue(this decimal val)
        {
            if (val == 0) return val;
            else
            {
                return val * -1;
            }
        }
        public static string PrimarySecondaryEntityFieldValToString(this Entity primaryEntity, Entity secondEntity, string primaryField, string secondField)
        {
            string ret = string.Empty;
            try
            {
                if (!string.IsNullOrEmpty(primaryEntity.EntityAttributeToString(primaryField)))
                    ret = primaryEntity.EntityAttributeToString(primaryField);
                else if (!string.IsNullOrEmpty(secondEntity.EntityAttributeToString(secondField)))
                    ret = secondEntity.EntityAttributeToString(secondField);
            }
            catch (Exception ex)
            {
            }
            return ret;
        }
        public static string ToFormattedMoneyString(this decimal val)
        {
            return string.Format(CultureInfo.InvariantCulture, "${0:#,#0}", val);
        }
        public static string ToFormattedMoneyString2Decimal(this decimal val)
        {
            return string.Format(CultureInfo.InvariantCulture, "${0:#,#0.00}", val);
        }
        public static string ToFormattedDateTimeString(this DateTime val)
        {
            return val.ToString("d MMM yyyy");
        }
        public static string DateTimeNowFetchFilter()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        }
        public static string DateTimeFetchFilter(this DateTime date)
        {
            return date.ToString("yyyy-MM-dd HH:mm:ss");
        }
        public static string DateFetchFilter(this DateTime date)
        {
            return date.ToString("yyyy-MM-dd");
        }
        #endregion

        #region General Save Helpers
        public static Guid CreateNewEntitySave(IOrganizationService service, Entity entity)
        {
            return service.Create(entity);
        }
        public static Guid CloneEntity(IOrganizationService service, Entity ExistingEntity, List<string> IgnorFields = null, List<KeyValuePair<string, object>> UpdatedFields = null)
        {
            Entity NewEntity = new Entity();
            NewEntity.LogicalName = ExistingEntity.LogicalName;
            foreach (KeyValuePair<string, object> attr in ExistingEntity.Attributes)
            {
                if (IgnorFields == null || !IgnorFields.Contains(attr.Key)) NewEntity.UpsertEntityAttribute(attr.Key, attr.Value);

            }
            if (UpdatedFields != null)
            {
                foreach (KeyValuePair<string, object> attr in UpdatedFields)
                {
                    NewEntity.UpsertEntityAttribute(attr.Key, UpdatedFields.Where(x => x.Key == attr.Key).Select(x => x.Value).FirstOrDefault());
                }
            }

            return service.Create(NewEntity);
        }
        public static void AssignOwnership(IOrganizationService service, EntityReference TargetER, EntityReference NewOwnerER)
        {
            AssignRequest assign = new AssignRequest
            {
                Assignee = NewOwnerER,
                Target = TargetER
            };
            service.Execute(assign);
        }
        public static void CreateNNRelationship(IOrganizationService service, EntityReference TargetER, EntityReference RelatedER, string RelationshipName)
        {
            AssociateRequest AssReq = new AssociateRequest
            {
                Target = TargetER,
                RelatedEntities = new EntityReferenceCollection
                {
                    RelatedER
                },
                Relationship = new Relationship(RelationshipName)
            };
            service.Execute(AssReq);
        }
        public static void CreateNNRelationshipBatch(IOrganizationService service, EntityReference TargetER, EntityReferenceCollection RelatedERCol, string RelationshipName)
        {
            AssociateRequest AssReq = new AssociateRequest
            {
                Target = TargetER,
                RelatedEntities = RelatedERCol,
                Relationship = new Relationship(RelationshipName)
            };
            service.Execute(AssReq);
        }
        public static void DeleteNNRelationship(IOrganizationService service, EntityReference TargetER, EntityReference RelatedER, string RelationshipName)
        {
            DisassociateRequest DissAssReq = new DisassociateRequest
            {
                Target = TargetER,
                RelatedEntities = new EntityReferenceCollection
                {
                    RelatedER
                },
                Relationship = new Relationship(RelationshipName)
            };
            service.Execute(DissAssReq);
        }
        public static void DeleteNNRelationshipBatch(IOrganizationService service, EntityReference TargetER, EntityReferenceCollection RelatedERCol, string RelationshipName)
        {
            DisassociateRequest DissAssReq = new DisassociateRequest
            {
                Target = TargetER,
                RelatedEntities = RelatedERCol,
                Relationship = new Relationship(RelationshipName)
            };
            service.Execute(DissAssReq);
        }
        public static void BulkCreate(IOrganizationService service, DataCollection<Entity> entities)
        {
            // Create an ExecuteMultipleRequest object.
            var multipleRequest = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses. 
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };

            // Add a CreateRequest for each entity to the request collection.
            foreach (var entity in entities)
            {
                CreateRequest createRequest = new CreateRequest { Target = entity };
                multipleRequest.Requests.Add(createRequest);
            }

            // Execute all the requests in the request collection using a single web method call.
            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);

        }
        public static void BulkUpdate(IOrganizationService service, DataCollection<Entity> entities)
        {
            // Create an ExecuteMultipleRequest object.
            var multipleRequest = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses. 
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };

            // Add a UpdateRequest for each entity to the request collection.
            foreach (var entity in entities)
            {
                UpdateRequest updateRequest = new UpdateRequest { Target = entity };
                multipleRequest.Requests.Add(updateRequest);
            }

            // Execute all the requests in the request collection using a single web method call.
            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);

        }
        public static void BulkDelete(IOrganizationService service, IEnumerable<EntityReference> entityReferences)
        {
            // Create an ExecuteMultipleRequest object.
            var multipleRequest = new ExecuteMultipleRequest()
            {
                // Assign settings that define execution behavior: continue on error, return responses. 
                Settings = new ExecuteMultipleSettings()
                {
                    ContinueOnError = false,
                    ReturnResponses = true
                },
                // Create an empty organization request collection.
                Requests = new OrganizationRequestCollection()
            };

            // Add a DeleteRequest for each entity to the request collection.
            foreach (var entityRef in entityReferences)
            {
                DeleteRequest deleteRequest = new DeleteRequest { Target = entityRef };
                multipleRequest.Requests.Add(deleteRequest);
            }

            // Execute all the requests in the request collection using a single web method call.
            ExecuteMultipleResponse multipleResponse = (ExecuteMultipleResponse)service.Execute(multipleRequest);
        }
        public static void SetEntityStatus(IOrganizationService service, EntityReference EF, int StateCode, int StatusCode)
        {
            try
            {
                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = EF,
                    State = new OptionSetValue(StateCode),
                    Status = new OptionSetValue(StatusCode)
                };
                service.Execute(setStateRequest);
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(string.Format("SetEntityStatus Exception: {0}", ex.ToString()));
            }
        }
        #endregion

        #region Money
        public static Money ToMoney(this decimal val)
        {
            Money ret = new Money();
            ret.Value = val;
            return ret;
        }
        public static Money ToMoney(this decimal? val)
        {
            Money ret = new Money();
            if (val != null && val.HasValue)
                ret.Value = val.Value;
            else
                ret.Value = 0;
            return ret;
        }
        public static Money ToMoney(this int val)
        {
            Money ret = new Money();
            ret.Value = val;
            return ret;
        }
        public static EntityReference ToEntityReference(this Guid val, string LogicalName)
        {
            EntityReference ret = new EntityReference();
            ret.Id = val;
            ret.LogicalName = LogicalName;
            return ret;
        }
        public static decimal ToPositiveValue(this decimal val)
        {
            if (val < 0) return val * -1;
            else return val;
        }
        public static int ToPositiveValue(this int val)
        {
            if (val < 0) return val * -1;
            else return val;
        }
        #endregion

        #region Simple Calculations
        public static decimal YearlyMoneyToMonthly(this Money val)
        {
            if (val != null && val.Value != 0)
                return val.Value / 12;
            else
                return 0;
        }

        public static decimal YearlyNullIntToMonthly(this int? val)
        {
            if (val != null && val.Value != 0)
                return val.Value / 12;
            else
                return 0;
        }

        public static double DoubleToPercentage(this double val)
        {
            return val / 100;
        }
        public static decimal DecimalToPercentage(this decimal val)
        {
            return val / 100;
        }
        public static decimal ToPercent(this decimal val)
        {
            if (val != 0)
                return val / 100;
            else
                return 0;
        }
        public static double ToPositiveNum(this double val)
        {
            if (val < 0) val = val * -1;
            return val;
        }
        public static double MoneyToDouble(this Money val)
        {
            double ret = 0;
            if (val != null)
                ret = Convert.ToDouble(val.Value);

            return ret;
        }
        public static decimal MoneyToDecimal(this Money val)
        {
            decimal ret = 0;
            if (val != null)
                ret = val.Value;

            return ret;
        }

        public static int ToOSVal(this OptionSetValue val)
        {
            int ret = 0;
            if (val != null)
                ret = val.Value;

            return ret;
        }

        #endregion

        #region Workflow Helpers
        #endregion

        #region Queue Helpers
        public static Guid AddItemToQueue(IOrganizationService _orgService, Guid QueueId, EntityReference TargetItemER)
        {
            Guid ret = Guid.Empty;
            try
            {
                AddToQueueRequest addRequest = new AddToQueueRequest
                {
                    Target = TargetItemER,
                    DestinationQueueId = QueueId
                };
                AddToQueueResponse res = _orgService.Execute(addRequest) as AddToQueueResponse;
                ret = res.QueueItemId;
            }
            catch (Exception ex)
            {
            }
            return ret;
        }
        #endregion

    }
}
