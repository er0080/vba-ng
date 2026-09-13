namespace VbaNg.Runtime.Library;

/// <summary>
/// The Financial module of the VBA standard library (MS-VBAL 6.1.2, Financial module). VBA
/// evaluates these formulas in the x87 extended format, rounding to Double only where a value
/// leaves the formula: the result, the operands and result of a power (the C runtime's pow on
/// Doubles, 1 + Rate included), the due factor 1 + Rate, and the objective Rate and IRR iterate
/// on. <see cref="Extended"/> reproduces that bit for bit (Financial golden, and about eight
/// hundred values probed from Excel; docs/vba-quirks.md under Financial). NPer at a nonzero rate
/// alone computes in Double, and IPmt and PPmt combine Pmt's and FV's Double results in Double.
/// </summary>
public static class Financial
{
    private const double IterationStep = 0.00001;
    private const double IterationEpsilon = 0.0000001;
    private const int MaxIterations = 40;

    /// <summary>Pmt(Rate, NPer, PV, [FV = 0], [Due = 0]); a zero NPer raises 5.</summary>
    public static double Pmt(double rate, double nper, double pv, double fv = 0, double due = 0)
    {
        if (nper == 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (rate == 0)
        {
            return (((Extended)(-fv) - pv) / nper).ToDouble();
        }

        var compound = Compound(rate, nper);
        return (-((fv + pv * compound) / (DueFactor(rate, due) * (compound - 1) / rate))).ToDouble();
    }

    /// <summary>FV(Rate, NPer, Pmt, [PV = 0], [Due = 0]).</summary>
    public static double FV(double rate, double nper, double pmt, double pv = 0, double due = 0)
    {
        if (rate == 0)
        {
            return ((Extended)(-pv) - (Extended)pmt * nper).ToDouble();
        }

        var compound = Compound(rate, nper);
        return ((Extended)(-pv) * compound - pmt * DueFactor(rate, due) * (compound - 1) / rate).ToDouble();
    }

    /// <summary>PV(Rate, NPer, Pmt, [FV = 0], [Due = 0]).</summary>
    public static double PV(double rate, double nper, double pmt, double fv = 0, double due = 0)
    {
        if (rate == 0)
        {
            return ((Extended)(-fv) - (Extended)pmt * nper).ToDouble();
        }

        var compound = Compound(rate, nper);
        return (-((fv + pmt * DueFactor(rate, due) * (compound - 1) / rate) / compound)).ToDouble();
    }

    /// <summary>NPer(Rate, Pmt, PV, [FV = 0], [Due = 0]); an unreachable target raises 5.</summary>
    public static double NPer(double rate, double pmt, double pv, double fv = 0, double due = 0)
    {
        if (rate <= -1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (rate == 0)
        {
            if (pmt == 0)
            {
                throw VbaErrors.InvalidProcedureCall();
            }

            return (-(((Extended)pv + fv) / pmt)).ToDouble();
        }

        // In Double, unlike the rest of the module (its x87 forms miss two of sixteen probes); the payment over the rate joins each side before the logarithms: the last bit of NPer(0.05 / 12, -1000, 200000) depends on it (Financial golden).
        var scaled = due != 0 ? pmt * (1 + rate) / rate : pmt / rate;
        var numerator = -fv + scaled;
        var denominator = pv + scaled;
        if (numerator < 0 && denominator < 0)
        {
            numerator = -numerator;
            denominator = -denominator;
        }
        else if (numerator <= 0 || denominator <= 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return (Math.Log(numerator) - Math.Log(denominator)) / Math.Log(rate + 1);
    }

    /// <summary>IPmt(Rate, Per, NPer, PV, [FV = 0], [Due = 0]); a period outside 1 to NPer raises 5.</summary>
    public static double IPmt(double rate, double per, double nper, double pv, double fv = 0, double due = 0)
    {
        if (per <= 0 || per >= nper + 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (due != 0 && per == 1)
        {
            return 0;
        }

        // With payments at the start of the period, the balance the interest is taken on is the future value two periods back of the principal less one payment (Financial golden, IPmt(0.05 / 12, 12, 360, 200000, 1000, 1)).
        var payment = Pmt(rate, nper, pv, fv, due);
        var elapsed = per - 1;
        if (due != 0)
        {
            pv += payment;
            elapsed = per - 2;
        }

        return FV(rate, elapsed, payment, pv, 0) * rate;
    }

    public static double PPmt(double rate, double per, double nper, double pv, double fv = 0, double due = 0) =>
        Pmt(rate, nper, pv, fv, due) - IPmt(rate, per, nper, pv, fv, due);

    /// <summary>DDB(Cost, Salvage, Life, Period, [Factor = 2]): double-declining balance, never depreciating below the salvage value.</summary>
    public static double DDB(double cost, double salvage, double life, double period, double factor = 2)
    {
        if (factor <= 0 || salvage < 0 || period <= 0 || period > life || life <= 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        if (cost <= 0)
        {
            return 0;
        }

        if (life < 2)
        {
            return cost - salvage;
        }

        if (life == 2)
        {
            return period > 1 ? 0 : cost - salvage;
        }

        if (period <= 1)
        {
            var first = ((Extended)cost * factor / life).ToDouble();
            var limit = cost - salvage;
            return first > limit ? limit : first;
        }

        // The period's depreciation, less the amount by which the total depreciation would pass cost minus salvage, the total taken as cost minus the remaining book value; that excess alone is figured in Double (Financial golden: DDB(10000, 1000, 5, 5) is 296 plus six ulps; Financial probe, the last period of a six-year life).
        var remaining = (((Extended)life - factor) / life).ToDouble();
        var depreciation = (Extended)factor * cost / life * Math.Pow(remaining, period - 1);
        var total = (cost - cost * (Extended)Math.Pow(remaining, period)).ToDouble();
        var excess = total - cost + salvage;
        if (excess > 0)
        {
            depreciation -= excess;
        }

        var result = depreciation.ToDouble();
        return result >= 0 ? result : 0;
    }

    /// <summary>SLN(Cost, Salvage, Life); a zero life raises 5.</summary>
    public static double SLN(double cost, double salvage, double life)
    {
        if (life == 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return (((Extended)cost - salvage) / life).ToDouble();
    }

    /// <summary>SYD(Cost, Salvage, Life, Period); a period outside 1 to Life raises 5.</summary>
    public static double SYD(double cost, double salvage, double life, double period)
    {
        if (salvage < 0 || period <= 0 || period > life)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return (((Extended)cost - salvage) * ((Extended)life - period + 1) * 2 / ((Extended)life * ((Extended)life + 1))).ToDouble();
    }

    /// <summary>NPV(Rate, ValueArray()); a rate of -1 or an empty array raises 5.</summary>
    public static double NPV(double rate, ReadOnlySpan<double> values)
    {
        if (rate == -1 || values.Length < 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        return PresentValue(values, rate).ToDouble();
    }

    /// <summary>IRR(ValueArray(), [Guess = 0.1]): the secant iteration VBA uses; no convergence raises 5.</summary>
    public static double IRR(ReadOnlySpan<double> values, double guess = 0.1)
    {
        if (guess <= -1 || values.Length <= 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var rate0 = guess;
        var npv0 = PresentValue(values, rate0).ToDouble();
        var rate1 = npv0 > 0 ? rate0 + IterationStep : rate0 - IterationStep;
        if (rate1 <= -1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var npv1 = PresentValue(values, rate1).ToDouble();
        for (var i = 0; i < MaxIterations; i++)
        {
            if (npv1 == npv0)
            {
                rate0 = rate1 > rate0 ? rate0 - IterationStep : rate0 + IterationStep;
                npv0 = PresentValue(values, rate0).ToDouble();
                if (npv1 == npv0)
                {
                    throw VbaErrors.InvalidProcedureCall();
                }
            }

            var rate = Secant(rate0, npv0, rate1, npv1);
            if (rate <= -1)
            {
                rate = (rate1 - 1) / 2;
            }

            var npv = PresentValue(values, rate).ToDouble();

            // IRR stops when the rate moves by less than the tolerance, whatever the scale of the flows (Financial probe: the same flows times a million and a thousandth); Rate stops on its objective.
            if (Math.Abs(rate - rate1) < IterationEpsilon)
            {
                return rate;
            }

            (rate0, npv0, rate1, npv1) = (rate1, npv1, rate, npv);
        }

        throw VbaErrors.InvalidProcedureCall();
    }

    /// <summary>MIRR(ValueArray(), FinanceRate, ReinvestRate); a rate of -1 raises 5, and flows without both signs raise 11.</summary>
    public static double MIRR(ReadOnlySpan<double> values, double financeRate, double reinvestRate)
    {
        if (financeRate == -1 || reinvestRate == -1 || values.Length <= 1)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var positive = SignedPresentValue(values, reinvestRate, positive: true);
        var negative = SignedPresentValue(values, financeRate, positive: false);
        if (negative.IsZero || positive.IsZero)
        {
            throw VbaErrors.DivisionByZero();
        }

        var count = values.Length;
        var ratio = (-positive * Math.Pow(reinvestRate + 1, count) / (negative * (Extended.One + financeRate))).ToDouble();
        return Math.Pow(ratio, 1.0 / (count - 1)) - 1;
    }

    /// <summary>Rate(NPer, Pmt, PV, [FV = 0], [Due = 0], [Guess = 0.1]): VBA's secant iteration; no convergence raises 5.</summary>
    public static double Rate(double nper, double pmt, double pv, double fv = 0, double due = 0, double guess = 0.1)
    {
        if (nper <= 0)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var rate0 = guess;
        var y0 = EvaluateRate(rate0, nper, pmt, pv, fv, due);
        var rate1 = y0 > 0 ? rate0 / 2 : rate0 * 2;
        var y1 = EvaluateRate(rate1, nper, pmt, pv, fv, due);
        for (var i = 0; i < MaxIterations; i++)
        {
            if (y1 == y0)
            {
                rate0 = rate1 > rate0 ? rate0 - IterationStep : rate0 + IterationStep;
                y0 = EvaluateRate(rate0, nper, pmt, pv, fv, due);
                if (y1 == y0)
                {
                    throw VbaErrors.InvalidProcedureCall();
                }
            }

            var rate = Secant(rate0, y0, rate1, y1);
            var y = EvaluateRate(rate, nper, pmt, pv, fv, due);
            if (Math.Abs(y) < IterationEpsilon)
            {
                return rate;
            }

            (rate0, y0, rate1, y1) = (rate1, y1, rate, y);
        }

        throw VbaErrors.InvalidProcedureCall();
    }

    // Rate's objective, the future value the payments leave, stored as a Double between steps.
    private static double EvaluateRate(double rate, double nper, double pmt, double pv, double fv, double due)
    {
        if (rate == 0)
        {
            return ((Extended)pv + (Extended)pmt * nper + fv).ToDouble();
        }

        var compound = Compound(rate, nper);
        return (pv * compound + pmt * DueFactor(rate, due) * (compound - 1) / rate + fv).ToDouble();
    }

    // The secant step, in x87 from the Doubles the objective was stored as (Financial probe: Rate matches under no other form).
    private static double Secant(double rate0, double y0, double rate1, double y1) =>
        ((Extended)rate1 - ((Extended)rate1 - rate0) * y1 / ((Extended)y1 - y0)).ToDouble();

    // The C runtime's pow on Doubles: 1 + Rate is rounded to a Double before it, and the power is a Double after it.
    private static Extended Compound(double rate, double nper) => Math.Pow(1 + rate, nper);

    // Payments at the start of each period scale by 1 + Rate rounded to a Double (Financial probe: every other form misses a payment-at-start case).
    private static Extended DueFactor(double rate, double due) => due != 0 ? 1 + rate : Extended.One;

    /// <summary>The net present value in x87, the discount factor multiplied forward and each flow divided by it, as VBA evaluates it (Financial golden).</summary>
    private static Extended PresentValue(ReadOnlySpan<double> values, double rate)
    {
        var divisor = Extended.One + rate;
        var factor = Extended.One;
        Extended total = 0.0;
        foreach (var value in values)
        {
            factor *= divisor;
            total += value / factor;
        }

        return total;
    }

    private static Extended SignedPresentValue(ReadOnlySpan<double> values, double rate, bool positive)
    {
        var divisor = Extended.One + rate;
        var factor = Extended.One;
        Extended total = 0.0;
        foreach (var value in values)
        {
            factor *= divisor;
            var flow = (positive && value > 0) || (!positive && value < 0) ? value : 0;
            total += flow / factor;
        }

        return total;
    }

    /// <summary>The Double elements of a Double() or Variant() argument; anything else raises 13, an unallocated array 5.</summary>
    public static double[] Flows(in Variant array)
    {
        if (!array.IsArray)
        {
            throw VbaErrors.TypeMismatch();
        }

        var source = array.AsArray();
        if (!source.IsAllocated)
        {
            throw VbaErrors.InvalidProcedureCall();
        }

        var values = new double[source.Count];
        var i = 0;
        for (var index = 0; index < source.Count; index++)
        {
            values[i++] = Coerce.ToDouble(source.ElementAt(index));
        }

        return values;
    }
}
