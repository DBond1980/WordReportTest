using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VT_Test.DataModels
{
    public enum PrincipleEnum
    {
        InductiveSingle,
        InductiveThree,
        Capacitive,
        Lpvt
    }
    public enum InsulationEnum
    {
        SinglePole,
        DoublePole
    }
    public enum StandardEnum
    {
        Iec618693,
        Gost1983,
        Gost23625
    }
    public enum TestingProgramEnum
    {
        Short,
        Full
    }
}
