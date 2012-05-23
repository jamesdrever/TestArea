﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Umbraco.Forms.Core;
using Umbraco.Forms.Core.Attributes;
using Umbraco.Forms.Data;
using umbraco.cms.businesslogic.member;
using umbraco.cms.businesslogic.property;
using umbraco.cms.businesslogic.datatype;
using Umbraco.Forms.Data.Storage;
using MailChimp;
using MailChimp.Types;
using Umbraco.Forms.Core.Enums;

/// <summary>
/// Class creates a workflow for Umbraco Contour that registers
/// the user details in the Contour form
/// with MailChimp.  Uses the MCAPI.NET wrapper at http://mcapinet.codeplex.com/
/// Also needs a reference to FSharp.Core
/// Uses Mail Chimp merges: 
/// First Name (Contour form field "First Name" >> Mail Chimp field "FNAME"
/// Last Name  (Contour form field "Last Name" >> Mail Chimp field "LNAME"
/// </summary>
public class ContourMailChimpSubscription : WorkflowType
{

    /// <summary>
    /// Gets or sets the API key.
    /// </summary>
    /// <value>
    /// The API key.
    /// </value>
    [Setting("Mail Chimp API Key", description = "Mail Chimp API Key",
        control = "Umbraco.Forms.Core.FieldSetting.TextField") ]
    public string APIKey { get; set; }

    /// <summary>
    /// Gets or sets the list ID.
    /// </summary>
    /// <value>
    /// The list ID.
    /// </value>
    [Setting("Mail Chimp List ID", description = "Mail Chimp List ID",
        control = "Umbraco.Forms.Core.FieldSetting.TextField") ]
    public string ListID{ get; set; }

    /// <summary>
    /// Validates the settings.
    /// </summary>
    /// <returns></returns>
    public override List<Exception> ValidateSettings()
    {
        List<Exception> exceptions = new List<Exception>();

        return exceptions;
    }


    /// <summary>
    /// Initializes a new instance of the <see cref="ContourMailChimpSubscription"/> class.
    /// </summary>
    public ContourMailChimpSubscription()
    {
        this.Id = new Guid("605BA4FE-9517-445B-97C2-FD4B2B348EB2");
        this.Name = "Subscribe with Mail Chimp";
        this.Description = "Subscribes the user details with Mail Chimp";

    }

    /// <summary>
    /// Executes the specified record.
    /// </summary>
    /// <param name="record">The record.</param>
    /// <param name="e">The <see cref="Umbraco.Forms.Core.RecordEventArgs"/> instance containing the event data.</param>
    /// <returns>a WorkflowExecutionStatus.Completed is the subscription works </returns>
    public override Umbraco.Forms.Core.Enums.WorkflowExecutionStatus Execute(Record record, RecordEventArgs e)
    {
        string firstName = record.GetRecordField("First Name").ValuesAsString();
        string lastName = record.GetRecordField("Last Name").ValuesAsString();
        string email = record.GetRecordField("Email").ValuesAsString();
        MCApi api = new MCApi(APIKey, false);

        var subscribeOptions =
         new Opt<List.SubscribeOptions>(
             new List.SubscribeOptions
             {
                 SendWelcome = true,
                 UpdateExisting = true,
                 DoubleOptIn = false,
             });

        var merges =
            new Opt<List.Merges>(
            new List.Merges
					{
						{"FNAME", firstName},
                        {"LNAME",lastName}
					});

        bool subscribeWorked = api.ListSubscribe(ListID, email, merges, subscribeOptions);
        if (subscribeWorked)
            return WorkflowExecutionStatus.Completed;
        else
            return WorkflowExecutionStatus.Failed;
    }
}