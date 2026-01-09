import * as React from 'react';
import { useState } from 'react';
import AddExpense from './AddExpense';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface Trip {
  Id: number;
  Title: string;
  Participants: { Id: number; Title: string }[];
}

interface TripManagerProps {
  context: WebPartContext;
  siteUrl: string;
}

const TripManager: React.FC<TripManagerProps> = ({ context, siteUrl }) => {
  const [showAddExpense, setShowAddExpense] = useState(false);
  const [selectedTrip, setSelectedTrip] = useState<Trip | null>(null);

  const fetchExpenses = () => {
    // fetch expenses logic
    console.log('Fetching expenses...');
  };

  const handleTripCreated = (trip: { Id: number; Title: string; TotalPackage: number }) => {
    console.log('Trip created:', trip);
    // handle new trip logic if needed
  };

  return (
    <div>
      <button
        className="btn btn-primary"
        onClick={() => {
          // Example: select a trip before adding expense
          setSelectedTrip({ Id: 1, Title: 'Trip to Goa', Participants: [] });
          setShowAddExpense(true);
        }}
      >
        Add Expense
      </button>

      {showAddExpense && selectedTrip && (
        <AddExpense
          context={context}
          siteUrl={siteUrl}
          trip={selectedTrip}
          tripId={selectedTrip.Id}
          onSave={() => setShowAddExpense(false)}
          onCancel={() => setShowAddExpense(false)}
          onTripCreated={handleTripCreated}
          onExpenseSaved={fetchExpenses}
        />
      )}
    </div>
  );
};

export default TripManager;