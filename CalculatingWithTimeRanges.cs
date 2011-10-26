		/// <summary>
		/// Check if the shift start hour is in the time range
		/// </summary>
		public bool ShiftFallsInTimeRange()
		{
			DateTime	dtRangeStart, dtRangeEnd,
				dtScheduleStartHour, dtScheduleEndHour;

			dtScheduleStartHour	= Convert.ToDateTime( drScheduleRow[ "SCHEDULED_TIME_IN" ] );
			dtScheduleEndHour	= Convert.ToDateTime( drScheduleRow[ "SCHEDULED_TIME_OUT" ] );

			// if equation is for holiday pay, just check if the start of shift 
			// is in the holiday
			if( strEquationType.Equals( "HOL" ) )
			{
				dtRangeStart	= dtHolidayRangeStart;
				dtRangeEnd		= dtHolidayRangeEnd;

				// check if holiday is associated to a location
				if( !Convert.IsDBNull( this.drHolidayRow ) && drHolidayRow != null )
					if( !Convert.IsDBNull( this.drHolidayRow[ "LOC_ID" ] ) )
					{
						Decimal	decHolidayLocId	= Convert.ToDecimal( this.drHolidayRow[ "LOC_ID" ] );

						// if a location was specified, check if it is the same 
						// as in the schedule
						if( !decHolidayLocId.Equals( 
							Convert.ToDecimal( this.drScheduleRow[ "LOC_ID" ] ) ) )
							// the holiday location is different than the schedule's, 
							// so the shift is not in a holiday for this location
							return false;
					}

				// check if start of shift falls between holiday's begin and end time
				
				// if shift started before holiday started, do not apply holiday rate
				if( dtScheduleStartHour.CompareTo( dtRangeStart ) < 0 )
					return false;

				// if shift started after holiday ended, do not apply holiday rate
				if( dtScheduleStartHour.CompareTo( dtRangeEnd ) > 0 )
					return false;

				// if end of shift falls between holiday's begin and end time 
				// or if start of shift is before holiday's begin time, 
				// we still pay regular rate instead of holiday rate
			}

			// use SCHEDULED times instead of ACTUAL times, 
			// but ShiftLength will still be according to actual times

			// if there is no time range, returns true
			if( bPayPlanHasNoTimeRange )
				return true;

			// get the intersection of the equation time 
			// range applied to the schedule start day
			dtRangeStart	= new DateTime( dtScheduleStartHour.Year,
				dtScheduleStartHour.Month,
				dtScheduleStartHour.Day,
				dtEquationStartHour.Hour,
				dtEquationStartHour.Minute,
				dtEquationStartHour.Second );

			dtRangeEnd	= new DateTime( dtScheduleStartHour.Year,
				dtScheduleStartHour.Month,
				dtScheduleStartHour.Day,
				dtEquationEndHour.Hour,
				dtEquationEndHour.Minute,
				dtEquationEndHour.Second );

			// if end hour is less than start hour, fix it by adding a day
			if( dtRangeStart.CompareTo( dtRangeEnd ) > 0 )
				dtRangeEnd	= dtRangeEnd.AddDays( 1 );

			// check if start of shift falls between equation's begin and end time
			// compare scheduled start hour to dtRangeEnd excluding coincidence of hours
			// because the time range is implied to finish at X-1 hours 59 minutes 59 seconds
			if( dtScheduleStartHour.CompareTo( dtRangeStart ) >= 0 )
				if( dtScheduleStartHour.CompareTo( dtRangeEnd ) < 0 )
					return true;

			// if end hour is before the start hour, it means begin hour isn't in same day
			// so if the shift ends in the current day, the dtRangeBegin will have to be 
			// adjusted to make it as the day before the schedule day
			if( dtEquationEndHour.CompareTo( dtEquationStartHour ) < 0 )
			{
				// get the intersection of the equation time 
				// range applied to the schedule start day 
				// (this time subtract one day from begin date)
				dtRangeStart	= new DateTime( dtScheduleStartHour.Year,
					dtScheduleStartHour.Month,
					dtScheduleStartHour.Day,
					dtEquationStartHour.Hour,
					dtEquationStartHour.Minute,
					dtEquationStartHour.Second );

				dtRangeEnd	= new DateTime( dtScheduleStartHour.Year,
					dtScheduleStartHour.Month,
					dtScheduleStartHour.Day,
					dtEquationEndHour.Hour,
					dtEquationEndHour.Minute,
					dtEquationEndHour.Second );

				// if end hour is less than start hour, fix it by subtracting a day
				if( dtRangeStart.CompareTo( dtRangeEnd ) > 0 )
					dtRangeStart	= dtRangeStart.AddDays( -1 );

				// check if start of shift falls between equation's begin and end time
				// compare scheduled start hour to dtRangeEnd excluding coincidence of hours
				// because the time range is implied to finish at X-1 hours 59 minutes 59 seconds
				if( dtScheduleStartHour.CompareTo( dtRangeStart ) >= 0 )
					if( dtScheduleStartHour.CompareTo( dtRangeEnd ) < 0 )
						return true;
			}

			return false;
		}



		public Decimal GetHoursInTimeRange( DateTime dtStart, DateTime dtEnd, 
			TimeSpan tsBeginTimeRange, TimeSpan tsEndTimeRange, String strActualOrScheduled )
		{
			int	iDays		= Convert.ToInt32( Math.Ceiling( 
				dtEnd.Subtract( dtStart ).TotalDays ) ) + 1;
			int	iDayCount	= 0;
			int iCurrentTimingCd	= iScheduleTimingCd;
			Decimal		decHours	= 0, decTotalHours	= 0;
			DateTime	dtCurrentTimeSlotBegin, dtCurrentTimeSlotEnd, dtCurrent	= dtStart;
			while( iDayCount < iDays )
			{
				iDayCount++;

				// check if time range crossed from one day to the next
				if( tsBeginTimeRange.CompareTo( tsEndTimeRange ) > 0 
					&& iDayCount == 1 )
				{
					// if it crossed and it is the first iteration, 
					// we need to check "yesterday" first
					dtCurrent	= dtCurrent.AddDays( -1 );

					// adjust Timing Cd (1 = Sun, 2 = Mon, ... 7 = Sat)
					iCurrentTimingCd	= ( iCurrentTimingCd == 1 ? 7 : iCurrentTimingCd - 1 );

					// since we are going back one day, increase the number of days to check
					iDays ++;
				}

				// Loop through each day that the shift spanned, 
				// add hours that match the time range
				dtCurrentTimeSlotBegin	= new DateTime( dtCurrent.Year, 
					dtCurrent.Month, 
					dtCurrent.Day, 
					tsBeginTimeRange.Hours, 
					tsBeginTimeRange.Minutes,
					tsBeginTimeRange.Seconds );
				dtCurrentTimeSlotEnd	= new DateTime( dtCurrent.Year, 
					dtCurrent.Month, 
					dtCurrent.Day, 
					tsEndTimeRange.Hours, 
					tsEndTimeRange.Minutes,
					tsEndTimeRange.Seconds );

				// if end hour is earlier than the begin hour, 
				// it means same end hour is in the next day
				if( dtCurrentTimeSlotBegin.CompareTo( dtCurrentTimeSlotEnd ) > 0 )
					dtCurrentTimeSlotEnd	= dtCurrentTimeSlotEnd.AddDays( 1 );

				// if equation is applicable to the day in question, count the hours
				if( IsEquationApplicableToDay( iEquationTimingCd, iCurrentTimingCd ) )
				{

					// match hours worked to the time slot, 
					// so we count only hours worked during the equation's time slot
					if( dtStart.CompareTo( dtCurrentTimeSlotBegin ) > 0 )
						dtCurrentTimeSlotBegin	= dtStart;
					if( dtEnd.CompareTo( dtCurrentTimeSlotEnd ) < 0 )
						dtCurrentTimeSlotEnd	= dtEnd;

					// get number of hours in between
					decHours	= Convert.ToDecimal( dtCurrentTimeSlotEnd.Subtract( dtCurrentTimeSlotBegin ).TotalHours );

					// adjust for DST if applicable
					if( IsDaylightSavings( dtCurrentTimeSlotBegin, dtCurrentTimeSlotEnd ) ) 
					{
						if( strActualOrScheduled.Equals( "ACTUAL" ) )
							decHours	+= Convert.ToDecimal( drScheduleRow[ "ACTUAL_DAYLIGHT" ] );
						else
							decHours	+= Convert.ToDecimal( drScheduleRow[ "SCHEDULED_DAYLIGHT" ] );
					}

					// accumulate number of hours matched to the time range
					// only add if hours is positive
					if( decHours > 0m )
						decTotalHours	+= decHours;

				}

				// proceed to verify the next day in the shift
				dtCurrent	= dtCurrent.AddDays( 1 );

				// adjust Timing Cd (1 = Sun, 2 = Mon, ... 7 = Sat)
				iCurrentTimingCd	= ( iCurrentTimingCd == 7 ? 1 : iCurrentTimingCd + 1 );
			}

			return decTotalHours;
		}

