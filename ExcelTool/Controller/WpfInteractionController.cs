﻿using AddinX.Logging;
using ExcelTool.Manipulation;
using Prism.Events;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTool.Controller
{
    public class WpfInteractionController : IDisposable
    {
        private readonly IEventAggregator eventAgg;
        private readonly ExcelInteraction excelOperation;
        private readonly ILogger logger;
        private readonly SubscriptionToken tokenMeetingData;
        private readonly SubscriptionToken tokenSheetName;

        public WpfInteractionController(IEventAggregator eventAgg,
            ExcelInteraction excelOperation,
            ILogger logger)
        {
            this.eventAgg = eventAgg;
            this.excelOperation = excelOperation;
            this.logger = logger;

            //tokenMeetingData = eventAgg.GetEvent<PubSubEvent<ExcelMeetingDataRequest>>()
            //    .Subscribe(WriteMeetingData);

            //tokenSheetName = eventAgg.GetEvent<PubSubEvent<ExcelWorksheetNamesRequest>>()
            //    .Subscribe(GetWorksheetsName);
        }

        //private void GetWorksheetsName(ExcelWorksheetNamesRequest obj)
        //{
        //    logger.Debug("Return excel sheets names");

        //    var response = new ExcelWorksheetNamesResponse
        //    {
        //        SheetNames = excelOperation.WorksheetsName().ToArray(),
        //    };

        //    eventAgg.GetEvent<PubSubEvent<ExcelWorksheetNamesResponse>>()
        //        .Publish(response);
        //}

        //private void WriteMeetingData(ExcelMeetingDataRequest obj)
        //{
        //    logger.Debug("Write Meeting data");
        //    try
        //    {
        //        excelOperation.WriteMeeting(obj);
        //        eventAgg.GetEvent<PubSubEvent<ExcelMeetingDataResponse>>()
        //            .Publish(new ExcelMeetingDataResponse
        //            { ProcessCompletedSuccessfully = true });
        //    }
        //    catch (Exception)
        //    {
        //        eventAgg.GetEvent<PubSubEvent<ExcelMeetingDataResponse>>()
        //            .Publish(new ExcelMeetingDataResponse
        //            { ProcessCompletedSuccessfully = false });
        //    }

        //}

        public void Dispose()
        {
        //    eventAgg.GetEvent<PubSubEvent<ExcelMeetingDataRequest>>()
        //        .Unsubscribe(tokenMeetingData);

        //    eventAgg.GetEvent<PubSubEvent<ExcelWorksheetNamesRequest>>()
        //        .Unsubscribe(tokenSheetName);
        }
    }
}
