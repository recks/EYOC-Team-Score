﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EYOC_Team_Score.Model
{
    public class PersonOrTeam {
        public required int Id { get; set; }
        public required string Name { get; set; }
        public required int Place { get; set; }
        public required int Time { get; set; }  // Time in seconds
        public required int Score { get; set; }
        public required string Country { get; set; }
    }
}
