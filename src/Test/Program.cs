using System;
using System.Collections.Generic;
using System.Text;
using ExcelToObject;
using System.IO;

namespace Test
{
	class Program
	{
		static void Main(string[] args)
		{
			var data = new GunGameData();
			data.LoadFromExcel(@"..\..\test.xlsx");
		}
	}

	public enum StageType
	{
		Stage_Easy,
		Stage_Normal,
		Stage_Hard,
	}

	public enum EnemyType
	{
		Speed,
		Balance,
		Tanker
	}

	public enum PlayerIncreaseCategory
	{
		Attack,
		Defence,
		Support
	}

	public partial class GunGameData
	{
		public void LoadFromExcel(string file)
		{
			var excelReader = new ExcelReader(file);

			StagePhase = excelReader.ReadList<StagePhaseData>("StageData");

			StageModify = excelReader.ReadDictionary<StageType, StageModifyData>("StageModifyData");

			SpawnCountRatio = excelReader.ReadDictionary<int, SpawnCountRatioData>("SpawnCountRatioData");

			Enemy = excelReader.ReadList<EnemyData>("EnemyData");

			Npc = excelReader.ReadDictionary<StageType, NpcData>("NpcData", x => x.matchStageType);

			Gun = excelReader.ReadList<GunData>("GunData");

			PointExchangeRatio = excelReader.ReadValue<float>("PointExchangeData", "exchangeRatio");

			PlayerIncrease = excelReader.ReadList<PlayerIncreaseData>("PlayerIncreaseData");

			CriticalDamageRate = excelReader.ReadValue<int>("CriticalDamageData", "criticalDamageRate");

			ChainKillBonus = excelReader.ReadList<ChainKillBonusData>("ChainKillsData");
		}

		public class StagePhaseData
		{
			public StageType type;
			public float duration;
			public float spawnInterval;
			public int maxCount;
			public float ratio_Npc;
			public Dictionary<EnemyType, float> ratio;
		}

		[ExcelToObject(TableName: "abc", SheetName: "sheet")]
		public List<StagePhaseData> StagePhase;

		public class StageModifyData
		{
			public StageType type;
			public float modify_Interval;
			public float spawnInterval_Modify;
		}
		public Dictionary<StageType, StageModifyData> StageModify;

		public class SpawnCountRatioData
		{
			public int maxCount;
			public float[] ratio;
		}
		public Dictionary<int, SpawnCountRatioData> SpawnCountRatio;

		public class EnemyData
		{
			public string id;
			public EnemyType type;
			public StageType matchStageType;
			public int hp;
			public float atkSpeed;
			public int gainPoint;
			public float atkSpeed_Modify;
			public string color;
		}
		public List<EnemyData> Enemy;

		public class NpcData
		{
			public string id;
			public StageType matchStageType;
			public float respawnCooltime;
			public float despawnTime;
			public int penaltyPoint;
		}
		public Dictionary<StageType, NpcData> Npc;

		public class GunData
		{
			public string id;
			public string name;
			public int atk;
			public int criticalRate;
			public int gainGoldBonus;
			public int price;
		}
		public List<GunData> Gun;

		public float PointExchangeRatio;

		public class PlayerIncreaseData
		{
			public string id;
			public PlayerIncreaseCategory category;
			public int level;
			public int addAtk;
			public int addCriticalRate;
			public int addHeart;
			public int addExemption;
			public int addGainGoldBonus;
			public int price;
		}
		List<PlayerIncreaseData> PlayerIncrease;

		public int CriticalDamageRate;

		public class ChainKillBonusData
		{
			public int startCount;
			public int bonusPointRate;
		}
		public List<ChainKillBonusData> ChainKillBonus;
	}
}

