import * as React from 'react';
import { DatePicker, DayOfWeek, Stack, Label, Text, Dropdown, IDropdownOption, IStackTokens, mergeStyles, Icon } from '@fluentui/react';

export interface DateTimePickerFieldProps {
  label?: string;
  value?: Date | string;
  onChange: (date?: Date) => void;
  required?: boolean;
  placeholder?: string;
}

export const DateTimePickerField: React.FC<DateTimePickerFieldProps> = ({
  label,
  value,
  onChange,
  required = false,
  placeholder = "Select date and time"
}) => {
  const [selectedDate, setSelectedDate] = React.useState<Date | undefined>(
    value ? (typeof value === 'string' ? (value ? new Date(value as string) : undefined) : value as Date) : undefined
  );

  // Generate time options in 15-minute increments
  const timeOptions: IDropdownOption[] = React.useMemo(() => {
    const options: IDropdownOption[] = [];
    for (let hour = 0; hour < 24; hour++) {
      for (let minute = 0; minute < 60; minute += 15) {
        const timeStr = `${('0' + hour).slice(-2)}:${('0' + minute).slice(-2)}`;
        const displayTime = hour === 0 && minute === 0 ? '12:00 AM' :
                           hour < 12 ? `${hour === 0 ? 12 : hour}:${('0' + minute).slice(-2)} AM` :
                           hour === 12 ? `12:${('0' + minute).slice(-2)} PM` :
                           `${hour - 12}:${('0' + minute).slice(-2)} PM`;
        options.push({
          key: timeStr,
          text: displayTime
        });
      }
    }
    return options;
  }, []);

  const containerStyles = mergeStyles({
    border: '1px solid #d2d0ce',
    borderRadius: '2px',
    padding: '8px',
    backgroundColor: '#ffffff',
    transition: 'border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out',
    ':hover': {
      borderColor: '#106ebe',
      boxShadow: '0 0 0 1px #106ebe'
    },
    ':focus-within': {
      borderColor: '#0078d4',
      boxShadow: '0 0 0 1px #0078d4'
    }
  });

  const stackTokens: IStackTokens = { childrenGap: 6 };

  const handleDateChange = (date?: Date | null) => {
    if (date) {
      const newDate = selectedDate ? new Date(selectedDate.getTime()) : new Date();
      newDate.setFullYear(date.getFullYear(), date.getMonth(), date.getDate());
      // Round minutes to nearest 15-minute increment if time hasn't been set
      if (!selectedDate) {
        const minutes = Math.round(newDate.getMinutes() / 15) * 15;
        newDate.setMinutes(minutes, 0, 0);
      }
      setSelectedDate(newDate);
      onChange(newDate);
    } else {
      setSelectedDate(undefined);
      onChange(undefined);
    }
  };

  const handleTimeChange = (event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption) => {
    if (option && typeof option.key === 'string') {
      const timeStr = option.key;
      const [hours, minutes] = timeStr.split(':').map(Number);
      
      if (selectedDate) {
        const newDate = new Date(selectedDate.getTime());
        newDate.setHours(hours, minutes, 0, 0);
        setSelectedDate(newDate);
        onChange(newDate);
      } else {
        // If no date selected but time is, use today's date
        const today = new Date();
        today.setHours(hours, minutes, 0, 0);
        setSelectedDate(today);
        onChange(today);
      }
    }
  };

  const clearDateTime = () => {
    setSelectedDate(undefined);
    onChange(undefined);
  };

  return (
    <Stack tokens={{ childrenGap: 6 }}>
      {label && (
        <Label required={required} styles={{ root: { fontWeight: 600, fontSize: '14px', color: '#323130' } }}>
          {label}
        </Label>
      )}
      <div className={containerStyles}>
        <Stack horizontal verticalAlign="center" tokens={stackTokens}>
          <Stack horizontal tokens={{ childrenGap: 8 }} styles={{ root: { flex: 1 } }}>
            <DatePicker
              placeholder="Select date"
              value={selectedDate}
              onSelectDate={handleDateChange}
              firstDayOfWeek={DayOfWeek.Sunday}
              showMonthPickerAsOverlay
              formatDate={(date?: Date) => date ? date.toLocaleDateString('en-US', { 
                month: 'short', 
                day: 'numeric', 
                year: 'numeric' 
              }) : ''}
              styles={{
                root: { flex: 1, minWidth: '160px' },
                textField: { 
                  border: 'none',
                  backgroundColor: 'transparent',
                  fontSize: '14px',
                  ':focus': {
                    border: 'none',
                    outline: 'none'
                  }
                },
                icon: {
                  color: '#0078d4',
                  fontSize: '16px'
                }
              }}
            />
            <Dropdown
              placeholder="Time"
              options={timeOptions}
              selectedKey={selectedDate ? `${('0' + selectedDate.getHours()).slice(-2)}:${('0' + selectedDate.getMinutes()).slice(-2)}` : undefined}
              onChange={handleTimeChange}
              styles={{
                root: { minWidth: '120px' },
                dropdown: { 
                  border: 'none',
                  backgroundColor: 'transparent',
                  fontSize: '14px',
                  ':focus': {
                    border: 'none',
                    outline: 'none'
                  }
                },
                caretDown: {
                  color: '#0078d4',
                  fontSize: '12px'
                }
              }}
            />
          </Stack>
          {selectedDate && (
            <Icon
              iconName="Cancel"
              styles={{
                root: {
                  fontSize: '12px',
                  color: '#a19f9d',
                  cursor: 'pointer',
                  padding: '4px',
                  ':hover': {
                    color: '#d13438'
                  }
                }
              }}
              onClick={clearDateTime}
              title="Clear date and time"
            />
          )}
        </Stack>
      </div>
      {selectedDate && (
        <Text variant="small" styles={{ root: { color: '#605e5c', fontStyle: 'italic', marginTop: '4px' } }}>
          <Icon iconName="Calendar" styles={{ root: { marginRight: '4px', fontSize: '12px' } }} />
          {selectedDate.toLocaleString('en-US', {
            weekday: 'short',
            year: 'numeric',
            month: 'short',
            day: 'numeric',
            hour: 'numeric',
            minute: '2-digit',
            hour12: true
          })}
        </Text>
      )}
    </Stack>
  );
};
