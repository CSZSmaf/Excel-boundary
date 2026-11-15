using System;
using System.Diagnostics;

// All timestamps use the monotonic stopwatch clock, including on long-running Excel sessions.
sealed class MaintenanceSchedule {
    long motionUntil, nextEntry, nextGeometry;
    uint geometryInput;
    public static long Now { get { return Stopwatch.GetTimestamp() * 1000 / Stopwatch.Frequency; } }
    public void ObserveMotion(long now) { motionUntil = now + 450; }
    public bool Settled(long now) { return now >= motionUntil; }
    public bool EntryDue(long now) { return now >= nextEntry; }
    public void EntryDone(long now) { nextEntry = now + 1000; }
    public bool GeometryDue(long now, uint idle, uint input, bool dirty) {
        return idle >= 1800 && now >= nextGeometry && (dirty || input != geometryInput);
    }
    public void GeometryDone(long now, uint input) { nextGeometry = now + 5000; geometryInput = input; }
}
