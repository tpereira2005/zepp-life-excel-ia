package core

import (
	"time"
)

type Weight struct {
	// main data
	Date   time.Time `json:"Date"`
	Weight float32   `json:"Weight"` // kg

	// almost all scales
	BMI       float32 `json:"BMI,omitempty"`
	BodyFat   float32 `json:"BodyFat,omitempty"`   // percent
	BodyWater float32 `json:"BodyWater,omitempty"` // percent
	BoneMass  float32 `json:"BoneMass,omitempty"`  // kg

	// some scales
	MetabolicAge   int     `json:"MetabolicAge,omitempty"`   // years
	MuscleMass     float32 `json:"MuscleMass,omitempty"`     // kg
	PhysiqueRating int     `json:"PhysiqueRating,omitempty"` // 1-9
	ProteinMass    float32 `json:"ProteinMass,omitempty"`    // kg
	ProteinRatio   float32 `json:"ProteinRatio,omitempty"`   // percent
	VisceralFat    int     `json:"VisceralFat,omitempty"`    // 1-12 or 13-59

	// rare scales
	BasalMetabolism    int     `json:"BasalMetabolism,omitempty"`    // kcal, Basal Metabolic Rate (BMR)
	BodyScore          int     `json:"BodyScore,omitempty"`          // points
	HeartRate          int     `json:"HeartRate,omitempty"`          // beats per minute
	Height             float32 `json:"Height,omitempty"`             // cm
	SkeletalMuscleMass float32 `json:"SkeletalMuscleMass,omitempty"` // kg
	StandBodyWeight    float32 `json:"StandBodyWeight,omitempty"`    // kg, raw Zepp summary field
	Impedance          int     `json:"Impedance,omitempty"`          // ohm
	EncryptImpedance   string  `json:"EncryptImpedance,omitempty"`   // raw Zepp summary field

	User   string `json:"User,omitempty"`
	Source string `json:"Source,omitempty"`
	ProfileID *int64 `json:"ProfileID,omitempty"` // Zepp member endpoint; -1 is the account owner.

	// unknown data
	//CaloricIntake int // ? Garmin

	// other
	//WHR float32 // Waist-to-Hip Ratio (WHR)
}

func Equal(w1, w2 *Weight) bool {
	return w1.Weight == w2.Weight &&
		w1.BMI == w2.BMI &&
		w1.BodyFat == w2.BodyFat &&
		w1.BodyWater == w2.BodyWater &&
		w1.BoneMass == w2.BoneMass &&
		w1.MetabolicAge == w2.MetabolicAge &&
		w1.MuscleMass == w2.MuscleMass &&
		w1.PhysiqueRating == w2.PhysiqueRating &&
		w1.ProteinMass == w2.ProteinMass &&
		w1.ProteinRatio == w2.ProteinRatio &&
		w1.VisceralFat == w2.VisceralFat &&
		w1.BasalMetabolism == w2.BasalMetabolism &&
		w1.BodyScore == w2.BodyScore &&
		w1.HeartRate == w2.HeartRate &&
		w1.Height == w2.Height &&
		w1.SkeletalMuscleMass == w2.SkeletalMuscleMass &&
		w1.StandBodyWeight == w2.StandBodyWeight &&
		w1.Impedance == w2.Impedance &&
		w1.EncryptImpedance == w2.EncryptImpedance
}
