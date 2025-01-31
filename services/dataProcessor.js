// dataProcessor.js
export function processFrontEndData(reqBody) {
    const { month, ...data } = reqBody; // Extract "month", rest goes into "data"

    const totalDays = Object.keys(data)
        .filter(key => key.startsWith('date-')) // Get only date keys
        .length;

    // Restructure into an array of objects
    const days = [];
    for (let day = 1; day <= totalDays; day++) {
        days.push({
            date: data[`date-${day}`],
            weekday: data[`weekday-${day}`],
            onTime: data[`on-time-${day}`] ,
            offTime: data[`off-time-${day}`] 
        });
    }

    return { month, days };
}