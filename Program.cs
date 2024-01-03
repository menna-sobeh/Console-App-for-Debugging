using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Services;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using LinkDev.Crm.Cs.Task;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;


//Created on - Succeeded on
namespace testDebugging
{
    public class Program
    {
        public static double WorkingDuration = 0;
        public static TimeSpan WorkingStartSpan = new TimeSpan();
        public static TimeSpan WorkingEndSpan = new TimeSpan();
        public static int Duration = 0;
        public static int totalMin = 0;

        

        public static IOrganizationService OrganizationService;
        static void Main(string[] args)
        {
            
        CalculateSLADuration obj = new CalculateSLADuration();
            CRMConnect(@"MBSDC\crmadmin", "linkP@ss", "https://misa-dev.linkdev.com/XRMServices/2011/Organization.svc");
            var x = OrganizationService;
            testFunction(x);
        }
        public static void CRMConnect(string UserName, string Password, string SoapOrgServiceUri)
        {
            try
            {
                ClientCredentials credentials = new ClientCredentials();
                credentials.UserName.UserName = UserName;
                credentials.UserName.Password = Password;
                Uri serviceUri = new Uri(SoapOrgServiceUri);
                OrganizationServiceProxy proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);



                proxy.EnableProxyTypes();
                OrganizationService = (IOrganizationService)proxy;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error while connecting to CRM " + ex.Message);
                Console.ReadKey();
            }
        }
    
        public static void testFunction(IOrganizationService service)
        {
           // DateTime startDate = new DateTime(2023, 10, 17, 10, 0, 0);
           // DateTime endDate = new DateTime(2023, 10, 26, 14, 0, 0);
            
            //get task record
            Guid entityId = new Guid("7cdbf4fa-ff89-ee11-9991-0022488b2ff9");
            //95ff752a-4289-ee11-998f-0022488b2ff9

            //Task Entity
            Entity task = service.Retrieve("task", entityId, new ColumnSet(true));
            var TaskCreatedonDate = task.GetAttributeValue<DateTime>("createdon");
            Console.WriteLine($"Task createdon:  {TaskCreatedonDate}");
            #region Print
            //Console.WriteLine($"\n\t\tTask Entity\n");
            //foreach (var x in task.Attributes)
            //    Console.WriteLine("Key: {0}, \t\tValue: {1}", x.Key, x.Value); 
            #endregion
            //slaskpiInstance  //first response by kpi  //lookup
            EntityReference slakpiInstance = task.GetAttributeValue<EntityReference>("ldv_firstresponsebykpi"); //lookup

           //slakpiInstance [all]
            var slakpi = service.Retrieve(slakpiInstance.LogicalName, slakpiInstance.Id, new ColumnSet(true));
            var SLAKPISucceeded = slakpi.GetAttributeValue<DateTime>("succeededon");
            Console.WriteLine($"SLASucceeded on:  {SLAKPISucceeded}");


            #region printing
            //Console.WriteLine("\n\t\tSLAKPIInstance\n");
            //foreach (var x in slakpi.Attributes)
            //    Console.WriteLine("Key: {0}, \t\tValue: {1}", x.Key, x.Value); 
            #endregion

            //SLA   //lookup
            EntityReference slaRef = task.GetAttributeValue<EntityReference>("slaid");

            //SLA  //Entity
            Entity sla = service.Retrieve(slaRef.LogicalName, slaRef.Id, new ColumnSet(true));

            
            //Calendar
            EntityReference businessHoursRef = sla.GetAttributeValue<EntityReference>("businesshoursid"); //default calendar lookup
           
            if (businessHoursRef == null)
            {
                Console.WriteLine("Nooooooooooooooooo business Hours Ref");

                var startDate = TaskCreatedonDate;
                var endDate = SLAKPISucceeded;

                TimeSpan timeDiff = endDate.Subtract(startDate);
                //Same Day
                if (startDate.Date == endDate.Date)
                {
                    totalMin = Convert.ToInt32((timeDiff.TotalMinutes ));
                    Duration = totalMin;
                    Console.WriteLine("total min in same day:" + totalMin.ToString());
                    Console.WriteLine("total min in same day:" + Duration.ToString());
                }
                else //Diff Days
                {
                    Duration = Convert.ToInt32((timeDiff.Days * 24 * 60) + (timeDiff.Hours * 60) + (timeDiff.Minutes));
                }

                Console.WriteLine($"{timeDiff.Days}:{timeDiff.Hours}:{timeDiff.Minutes}");
                Console.WriteLine("Durationnnnnn:"+Duration);
            }
            else
            {

                //the schema name of the Business Hours entity is calendar 
                Entity businessCalendarObject = service.Retrieve("calendar", businessHoursRef.Id, new ColumnSet(true));

                //3- Get Vacations
                List<DateTime> vacationList = GetVacations(businessCalendarObject, service);

                //4- Working Days Pattern
                List<string> workingDaysPatternList = GetWorkingDaysPattern(businessCalendarObject);

                //  GetWorkingDaysPattern(businessCalendarObject);
                var holidaysCalendarEntity = businessCalendarObject.GetAttributeValue<EntityReference>("holidayschedulecalendarid");
                Entity Holiday = service.Retrieve(holidaysCalendarEntity.LogicalName, holidaysCalendarEntity.Id, new ColumnSet(true));

                // get calendar rules //21 att
                // var calendarRules2 = Holiday.GetAttributeValue<EntityCollection>("calendarrules");

                #region print
                //Console.WriteLine($"\nHoliday\n");
                //foreach (var temp in finalVacationList.OrderBy(x => x.Date).ToList())
                //{
                //    Console.WriteLine(temp);
                //} 
                #endregion


                //calendarrule Entities[0]
                var calendarRules = businessCalendarObject.GetAttributeValue<EntityCollection>("calendarrules");
                #region Printing
                // var calendarCollection = calendarRules.Entities[0].Attributes;
                //Console.WriteLine("\n\t\tCalendarrule\n");
                //foreach (var x in calendarCollection)
                //    Console.WriteLine("Key: {0}, \t\tValue: {1}", x.Key, x.Value); 
                #endregion

                // Get the ID of the inner calendar
                var innerCalendarId = calendarRules[0].GetAttributeValue<EntityReference>("innercalendarid").Id; //lookup

                // Retrieve the inner calendar with all of its columns //calendar Entity //12 att
                var innerCalendar = service.Retrieve("calendar", innerCalendarId, new ColumnSet(true));
                #region printing
                //Console.WriteLine("\n\t\tCalendar\n");
                //foreach (var x in innerCalendar.Attributes)
                //    Console.WriteLine("Key: {0}, \t\tValue: {1}", x.Key, x.Value); 
                #endregion

                #region toBeSeen
                //Calendar rule Entity //16 att // Get the first inner calendar rule
                // var innnerCalendarRule = innerCalendar.GetAttributeValue<EntityCollection>("calendarrules").Entities.FirstOrDefault();

                #endregion
                #region printing
                //Console.WriteLine("\n\t\tCalendarrule\n");
                //foreach (var x in innnerCalendarRule.Attributes)
                //    Console.WriteLine("Key: {0}, \t\tValue: {1}", x.Key, x.Value);
                #endregion
                if (businessCalendarObject != null && businessCalendarObject.Id != Guid.Empty)
                {
                    // Get the first inner calendar rule 
                    var innerCalendarRule = innerCalendar.GetAttributeValue<EntityCollection>("calendarrules").Entities.FirstOrDefault();

                    var startTime = Convert.ToDouble(innerCalendarRule.GetAttributeValue<int>("offset")) / Convert.ToDouble(60);
                    WorkingDuration = Convert.ToDouble(innerCalendarRule.GetAttributeValue<int>("duration")) / Convert.ToDouble(60);


                    WorkingStartSpan = TimeSpan.FromHours(startTime); //9
                    WorkingEndSpan = TimeSpan.FromHours(startTime + WorkingDuration); //17
                    Console.WriteLine($"WorkingStartSpanin: {WorkingStartSpan}");
                    Console.WriteLine($"WorkingEndSpanin: {WorkingEndSpan}");
                }

                

                //Get The Total Duration
                Duration = GetFinalDuration(TaskCreatedonDate, SLAKPISucceeded, workingDaysPatternList, vacationList);
                //     Duration = GetFinalDuration(startDate, endDate, workingDaysPatternList, vacationList);

                Console.WriteLine($"\n\tDuraion in Hours =  {Duration}\n");

               // throw new Exception("End of code");
            }
            Console.WriteLine("Final Duration = "+Duration);
        }


        public static int GetFinalDuration(DateTime startDate, DateTime endDate, List<string> workingDaysPatternList, List<DateTime> vacationList)
        {
           // tracingService.Trace("Inside GetFinalDuration");
            string durationInDays = string.Empty;
            string durationInHours = string.Empty;
            string durationInDaysandHours = string.Empty;
            List<DateTime> originalDatesRangeList = new List<DateTime>();

            #region Prepare Lists 
            #region comment

            // get range of dates before excluding the vacation and weekends
            //tracingService.Trace("1- Get Range of Dates");

            // case of the end date is less than the start date due to the creation and completion of the task are after working hours 
            // so the creation date is shiftted to the next working day and the completion date is still the same
            // so in this case, end date should == start date which means no time was taken 
            #endregion
            if ((endDate - startDate).Days < 0)
            {
                endDate = startDate;
                originalDatesRangeList.Add(startDate);
            }
            else
            {
                originalDatesRangeList = Enumerable.Range(0, (endDate - startDate).Days).Select(d => startDate.AddDays(d)).ToList();
                originalDatesRangeList.Add(endDate); //add endDate Seperate to keep the time as it is not as the above list date's time is the same as start date
            }

            //tracingService.Trace("2- Exclude Vacation");
            var datesWithoutVacationList = originalDatesRangeList.Where(originalDay => !vacationList.Any(vacationDay => vacationDay.Date == originalDay.Date));

            //tracingService.Trace("3- Exclude Weekends");  
            var dateWithoutOffDaysList = datesWithoutVacationList.Where(workingDay => workingDaysPatternList.Any(patternDay => patternDay.ToLower() == workingDay.DayOfWeek.ToString().Substring(0, 2).ToLower()));


            #endregion

            #region Calculate Total Hours
            //tracingService.Trace("4- Calculate Total Hours");
            var datesInBetween = dateWithoutOffDaysList.Where(workingDay => workingDay != startDate && workingDay != endDate);

            var totalHoursForDatesInBetween = (datesInBetween.Count() * WorkingDuration);

            TimeSpan totalHoursInStartDate = startDate.Hour > WorkingEndSpan.Hours ? TimeSpan.Zero :
                                    startDate.Hour < WorkingStartSpan.Hours ?
                                    WorkingEndSpan.Subtract(WorkingStartSpan) :
                                    WorkingEndSpan.Subtract(startDate.TimeOfDay);


            TimeSpan totalHoursInEndDate = endDate.Hour < WorkingStartSpan.Hours ?
                                      TimeSpan.Zero :
                                      endDate.Hour > WorkingEndSpan.Hours ?
                                      WorkingEndSpan.Subtract(WorkingStartSpan) :
                                      endDate.TimeOfDay.Subtract(WorkingStartSpan);

            Console.WriteLine($"--Total Hours in Start Date:{totalHoursInStartDate}");
            Console.WriteLine($"--Total Hours in End Date:{totalHoursInEndDate}");

            double totalHoursOfStartAndEndDates = 0;

            if (startDate.Date == endDate.Date)
                totalHoursOfStartAndEndDates = (endDate.AddSeconds(-endDate.Second) - startDate.AddSeconds(-startDate.Second)).TotalHours;
            else
                totalHoursOfStartAndEndDates = totalHoursInStartDate.TotalHours + totalHoursInEndDate.TotalHours;
            #endregion

            #region Calculate Days and hours and minutes seperates
            //tracingService.Trace("5- Calculate Days and hours and minutes seperates");

            double finalTotalInDays = Convert.ToDouble(datesInBetween.Count());

            double finalTotalHours = 0;
            double finalTotalMinutes = 0;

            if (totalHoursOfStartAndEndDates > WorkingDuration) //go here
            {

                var totalDiff = totalHoursOfStartAndEndDates / WorkingDuration;

                var daystoAdd = Math.Truncate(totalDiff);

                var hoursAndMins = (totalDiff - daystoAdd) * WorkingDuration;

                finalTotalHours = hoursAndMins > 0 ? Math.Truncate(hoursAndMins) : 0;

                finalTotalMinutes = (hoursAndMins - finalTotalHours) * 60;

                finalTotalInDays += daystoAdd;

            }
            else
            {
                finalTotalHours = Math.Truncate(totalHoursOfStartAndEndDates);

                finalTotalMinutes = (totalHoursOfStartAndEndDates - finalTotalHours) * 60;
            }

            #endregion

            #region Calculate Duration 
            //tracingService.Trace("6- Calculate Duration");

            durationInDays = dateWithoutOffDaysList.Count() > 0 ? dateWithoutOffDaysList.Count().ToString() : "1";

            durationInHours = String.Format("{0:0.00}", (totalHoursForDatesInBetween + totalHoursOfStartAndEndDates)).ToString();

            durationInDaysandHours = finalTotalInDays + " d, " + finalTotalHours + " h, " + Math.Truncate(finalTotalMinutes) + " m";
            Console.WriteLine(durationInDaysandHours);
            //total in minutes
            Duration = Convert.ToInt32((finalTotalInDays * 24 * 60) + (finalTotalHours * 60) + Math.Truncate(finalTotalMinutes));

             
           // Console.WriteLine("Duration in Minutes:"+testDuration.ToString());


            #endregion
            //tracingService.Trace("7- Finished");
            return Duration;
        //    return durationInHours;
        }
        public static List<DateTime> GetVacations(Entity businessCalendarObject, IOrganizationService OrganizationService)
        {
            List<DateTime> finalVacationList = new List<DateTime>();

            if (businessCalendarObject != null && businessCalendarObject.Id != Guid.Empty)
            {

                // check that the calendar object has a holiday schedule calendar linked to it
                if (businessCalendarObject.Contains("holidayschedulecalendarid") && businessCalendarObject.Attributes["holidayschedulecalendarid"] != null)
                {

                    var holidaysCalendarEntity = (EntityReference)businessCalendarObject.Attributes["holidayschedulecalendarid"];

                    // retrieve holidays object
                    var holidaysCalendarObject = OrganizationService.Retrieve(holidaysCalendarEntity.LogicalName, holidaysCalendarEntity.Id, new ColumnSet(true));

                    // get calendar rules
                    var calendarRules = holidaysCalendarObject.GetAttributeValue<EntityCollection>("calendarrules");

                    foreach (var calenderRule in calendarRules.Entities)
                    {
                        List<DateTime> vacationList = new List<DateTime>();

                        var startDate = (DateTime)calenderRule.Attributes["effectiveintervalstart"];
                        var endDate = (DateTime)calenderRule.Attributes["effectiveintervalend"];

                        // prepare the list of the vacation based on:
                        // 1- Total number of days between startDate and endDate "(endDate - startDate).Days"
                        // 2- Add total to the startDate "startDate.AddDays(d)" to generate the rest of the days

                        // i.e. Start 1/1(Jan)/2020 and End 3/1(Jan)/2020
                        // then, total days = (End - Start).Days + 1 = 3
                        // list will contains [1/1, 1/2, 1/3]
                        vacationList = Enumerable.Range(0, (endDate - startDate).Days).Select(d => startDate.AddDays(d)).ToList();

                        finalVacationList.AddRange(vacationList);
                    }
                }

            }


            return finalVacationList.OrderBy(x => x.Date).ToList();
        }
        public static List<string> GetWorkingDaysPattern(Entity businessCalendarObject)
        {

            List<string> patternDaysShortName = new List<string>();

            if (businessCalendarObject != null && businessCalendarObject.Id != Guid.Empty)
            {
                var calendarRules = businessCalendarObject.GetAttributeValue<EntityCollection>("calendarrules");

                var firstRulePattern = calendarRules[0].GetAttributeValue<string>("pattern");
                // FREQ=WEEKLY;INTERVAL=1;BYDAY=SU,MO,TU,WE,TH,FR

                patternDaysShortName = firstRulePattern.Substring(firstRulePattern.LastIndexOf("BYDAY")).Split(new string[] { "BYDAY=", "," }, StringSplitOptions.RemoveEmptyEntries).ToList();


            }


            return patternDaysShortName;
        }
   
    }
}
