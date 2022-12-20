using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using System.ServiceModel;
using Microsoft.Xrm.Sdk.Query;
using System.Security.Policy;
using System.IdentityModel.Metadata;
using System.Security.Principal;

namespace Prospect_OnCreate
{
    public class Class1 : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            ITracingService tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            if (context.InputParameters.Contains("Target") && context.InputParameters["Target"] is Entity)
            {
                Entity prospectEntity = (Entity)context.InputParameters["Target"];

                try
                {
                    int points = 0;
                    string name = (string)prospectEntity["fisoft_name"];
                    int age = (int)prospectEntity["fisoft_age"];
                    if (age >= 18)
                    {
                        points += 10; 
                    }
                    // high school = ,000 ; bachelor = ,001 ; master = ,002
                    OptionSetValue education = (OptionSetValue)prospectEntity["fisoft_education"];
                    int educationSelectedValue = education.Value;
                    switch (educationSelectedValue)
                    {
                        case 874630000:
                            points += 10;
                            break;

                        case 874630001:
                            points += 30;
                            break;

                        case 874630002:
                            points += 50;
                            break;

                        default:
                            points += 0;
                            break;
                    }

                    //engeenering = ,000 ; economics = ,001 ; law = ,002 ; medical = ,003
                    OptionSetValue profesion = (OptionSetValue)prospectEntity["fisoft_professionn"];
                    int profesionSelectedValue = education.Value;
                    switch (profesionSelectedValue)
                    {
                        case 874630000:
                            points += 50;
                            break;

                        case 874630001:
                            points += 40;
                            break;

                        case 874630002:
                            points += 30;
                            break;

                        case 874630003:
                            points += 10;
                            break;

                        default:
                            points += 0;
                            break;
                    }

                    // single = ,000 ; engaged = ,001 ; married = ,002 ; complicated = ,003
                    OptionSetValue relationshipStatus = (OptionSetValue)prospectEntity["fisoft_relationshipstatus"];
                    int relationshipStatusSelectedTopicValue = relationshipStatus.Value;
                    switch (relationshipStatusSelectedTopicValue)
                    {
                        case 874630000: 
                            points += 10;
                            break;

                        case 874630001:
                            points += 30;
                            break;

                        case 874630002:
                            points += 50;
                            break;
                        
                        case 874630003:
                            points += 0;
                            break;

                        default:
                            points += 0;
                            break;
                    }

                    if(points >= 100)
                    {
                        Entity newContact = new Entity("fisoft_contact");
                        newContact["fisoft_name"] = name;
                        newContact["fisoft_age"] = age;
                        newContact["fisoft_education"] = new OptionSetValue(educationSelectedValue) ;
                        newContact["fisoft_profession"] = new OptionSetValue(profesionSelectedValue) ;
                        newContact["fisoft_relationshipstatus"] = new OptionSetValue(relationshipStatusSelectedTopicValue);
                        service.Create(newContact);
                    }
                }
                catch (FaultException<OrganizationServiceFault> ex)
                {
                    throw new InvalidPluginExecutionException("An error occurred in FollowUpPlugin.", ex);
                }
            }
        }
    }
}
