using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FdxSNBCounter
{
    public sealed partial class GetUseronCounter : CodeActivity
    {
        [RequiredArgument]
        [Input("Counter")]
        [ReferenceTarget("fdx_snbcounter")]
        public InArgument<EntityReference> InputEntity { get; set; }

        [Output("user")]
        [ReferenceTarget("systemuser")]
        public OutArgument<EntityReference> OutputEntity { get; set; }

        [Output("NewCounter")]
        public OutArgument<int> OutputCounter { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            try
            {
                IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
                IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
                IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

                //Get entity from InArgument....
                Entity counter = service.Retrieve("fdx_snbcounter", this.InputEntity.Get(executionContext).Id, new ColumnSet("fdx_currentcounter"));

                if(counter.Contains("fdx_currentcounter"))
                {
                    //Fetch xml to get users of a team. Here one user will be returned per page and page position will be
                    //decided on current counter value.
                    string xmlQuery = string.Format("<fetch version='1.0' count='1' distinct='true' returntotalrecordcount='true' page='{0}' ><entity name='systemuser' ><attribute name='systemuserid' /><order attribute='fullname' descending='false' /><link-entity name='teammembership' from='systemuserid' to='systemuserid' visible='false' intersect='true' ><link-entity name='team' from='teamid' to='teamid' alias='aa' ><filter type='and' ><condition attribute='name' operator='eq' value='Sales Next@Bat' /></filter></link-entity></link-entity></entity></fetch>", (int)counter["fdx_currentcounter"]);

                    FetchExpression query = new FetchExpression(xmlQuery);
                    EntityCollection users = service.RetrieveMultiple(query);

                    //Set out argument OutputEntity- first user entity of resultset.
                    this.OutputEntity.Set(executionContext, new EntityReference("systemuser", new Guid(users[0].Attributes["systemuserid"].ToString())));

                    int newCounter = 1;
                    //Counter value should not exceed to total member count of team, in case if it exceeds
                    //that means user has been deleted from team.
                    if((int)counter["fdx_currentcounter"] >= users.TotalRecordCount)
                    {
                        //Set default value = 1
                        newCounter = 1;
                    }
                    else
                    {
                        //Increment by 1 current counter value
                        newCounter = (int)counter["fdx_currentcounter"] + 1;
                    }

                    //Set out argument OutputCounter- new counter value, that will be further
                    //used by workflow to update counter entity.
                    this.OutputCounter.Set(executionContext, newCounter);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException("GetUseronCounter ERROR>>>>>: " + ex.StackTrace.ToString(), ex);
            }
        }
    }
}
