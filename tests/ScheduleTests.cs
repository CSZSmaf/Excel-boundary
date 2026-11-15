using System;
class ScheduleTests {
    static void Check(bool value, string name) { if (!value) throw new Exception(name); }
    static int Main() {
        var schedule = new MaintenanceSchedule();
        long time = (long)Int32.MaxValue + 10000;
        schedule.ObserveMotion(time);
        Check(!schedule.Settled(time + 449), "wait for gesture settlement");
        Check(schedule.Settled(time + 450), "resume after gesture settlement");
        Check(schedule.EntryDue(time), "initial entry check");
        schedule.EntryDone(time);
        Check(!schedule.EntryDue(time + 999) && schedule.EntryDue(time + 1000), "one entry check per second");
        Check(!schedule.GeometryDue(time, 1799, 1, true), "geometry waits for idle input");
        Check(schedule.GeometryDue(time, 1800, 1, true), "geometry after idle input");
        schedule.GeometryDone(time, 1);
        Check(!schedule.GeometryDue(time + 4999, 2000, 2, true), "geometry rate limit");
        Check(!schedule.GeometryDue(time + 5000, 2000, 1, false), "unchanged idle work skipped");
        Check(schedule.GeometryDue(time + 5000, 2000, 2, false), "new input refreshes geometry");
        Console.WriteLine("Scheduler tests passed (including long-running-session timestamps).");
        return 0;
    }
}
