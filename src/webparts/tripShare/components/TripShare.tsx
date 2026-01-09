import * as React from 'react';
import { useState } from 'react';

import { ITripShareProps } from './ITripShareProps';

import TripDashboard from './TripDashboard';
import AddTrip from './AddTrip';
import TripDetails from './TripDetails'; 
import AddExpense from './AddExpense';
import Summary from './Summary';

type Screen = 
  | 'home' 
  | 'addTrip' 
  | 'tripDetails' 
  | 'addExpense' 
  | 'summary' 
  | 'tripClosed';

interface Participant {
  Title: string;
  Id: number;
  Email?: string;
}

interface Trip {
  Id: number;
  Title: string;
  TripName: string;
  StartDate: string;
  EndDate: string;
  Status: string;
  Participants: Participant[];
  TotalPackage?: number;
}

const TripShare: React.FC<ITripShareProps> = ({ context,
  description, isDarkTheme, environmentMessage, hasTeamsContext, userDisplayName
}) => {
  const [currentScreen, setCurrentScreen] = useState<Screen>('home');
  const [selectedTrip, setSelectedTrip] = useState<Trip | null>(null);

  const onTripSelect = (trip: Trip) => {
    setSelectedTrip(trip);
    setCurrentScreen('tripDetails');
  };

  const onAddTripClick = () => { 
    setCurrentScreen('addTrip');
  };

  const onTripCreated = (trip: Trip) => {
    setSelectedTrip(trip);
    setCurrentScreen('tripDetails');
  };

  const onAddExpenseClick = () => {
    setCurrentScreen('addExpense');
  };

  const onExpenseAdded = () => {
    setCurrentScreen('tripDetails');
  };

  const onViewSummaryClick = () => {
    setCurrentScreen('summary');
  };

  const onCloseTrip = () => {
    setCurrentScreen('tripClosed');
  };

  const onBackToHome = () => {
    setSelectedTrip(null);
    setCurrentScreen('home');
  };

  const onBackToTripDetails = () => {
    setCurrentScreen('tripDetails');
  };

  return (
    <div>
      {currentScreen === 'home' && (
        <TripDashboard
          context={context}
          onTripSelect={onTripSelect}
          onAddTrip={onAddTripClick}
        />
      )}

      {currentScreen === 'addTrip' && (
        <AddTrip
          context={context}
          onSave={onTripCreated}
          onCancel={onBackToHome} 
          onTripAdded={onTripCreated}  
          siteUrl={context.pageContext.web.absoluteUrl}  
          onBack={onBackToHome}        
        />
      )}

      {currentScreen === 'tripDetails' && selectedTrip && (
        <TripDetails
          context={context}
          trip={selectedTrip}
          onAddExpense={onAddExpenseClick}
          onViewSummary={onViewSummaryClick}
          onBack={onBackToHome}
        />
      )}

      {currentScreen === 'addExpense' && selectedTrip && (
        <AddExpense
          context={context}
          trip={selectedTrip}
          onSave={onExpenseAdded}
          onCancel={onBackToTripDetails}

          onExpenseSaved={onExpenseAdded}
          tripId={selectedTrip.Id}
          siteUrl={context.pageContext.web.absoluteUrl} onTripCreated={function (trip: { Id: number; Title: string; TotalPackage: number; }): void {
            throw new Error('Function not implemented.');
          } }        />
      )}

      {currentScreen === 'summary' && selectedTrip && (
        <Summary
          context={context}
          trip={selectedTrip}
          onCloseTrip={onCloseTrip}
          onBack={onBackToTripDetails}
        />
      )}

      {currentScreen === 'tripClosed' && selectedTrip && (
        <Summary
          context={context}
          trip={selectedTrip}
          readOnly={true}
          onBack={onBackToHome} 
          onCloseTrip={onCloseTrip}  
        />
      )}
    </div>
  );
};

export default TripShare;