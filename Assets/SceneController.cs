using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using OfficeOpenXml;
using System.IO;
using System;
using System.Linq;

[System.Serializable]
public class Card
{
	public string name;
	public int id;
	public int hp;
	public int spd;
	public Attack atk;
	public Defense defense;
	public Chart chart;
	public Skill skill;
	public int rebirthTo;
	public Card()
	{
		atk = new Attack();
		defense = new Defense();
		chart = new Chart();
		skill = new Skill();
	}
}

[System.Serializable]
public enum AttackType
{
	Single,
	Splash,
	Mass
}

[System.Serializable]
public enum DefenseType
{
	Deduction,
	Exceed,
	Deduction_Exceed
}

[System.Serializable]
public class Attack
{
	public AttackType type;
	public int count;
	public int value;
}

[System.Serializable]
public class Defense
{
	public DefenseType type;
	public int value1;
	public int value2;
}

[System.Serializable]
public class Chart
{
	public int win;
	public int all;
	public float per;
}

[System.Serializable]
public class Team
{
	public List<CardInBattle> cards ;
	public List<CardInBattle> cardinfield;

	public Team()
	{
		cards = new List<CardInBattle>();
		cardinfield = new List<CardInBattle>();
	}

	public void MoveToBattle()
	{
		if(cardinfield.Count <4)
		{
			foreach(CardInBattle c in cards)
			{
				if(c.currentHp > 0)
				{
					cardinfield.Add(c);
					if(cardinfield.Count>=4)
						return;
				}
			}
		}
	}

	public bool IsLose()
	{
		foreach(CardInBattle c in cardinfield)
		{
			if(c.currentHp > 0)
			return false;
		}
		return true;
	}

	public int GetNearestAttacked(int i=0)
	{
		if(i<0)
			i=3;
		while(cardinfield[i].currentHp <=0)
		{
			i--;
			if(i<0)
				i=3;
		}
		return i;	
		
	}

	public CardInBattle GetRandomCard()
	{
		int i = UnityEngine.Random.Range(0,4);
		while(cardinfield[i].currentHp <=0)
		{
			i = UnityEngine.Random.Range(0,4);
		}
		return cardinfield[i];
	}

	public CardInBattle GetWeakestCard()
	{
		int pos = 0;
		int weakesthp =cardinfield[pos].currentHp;
		for(int i = 1; i < 4; i ++)
		{
			if(weakesthp < cardinfield[i].currentHp)
			{
				pos = i;
				weakesthp = cardinfield[pos].currentHp;
			}
		}

		return cardinfield[pos];
	}

}

[System.Serializable]
public enum TeamEnum
{
	Attack= 0,
	Defense
}

#region skill class
public class BaseSkill
{
	
}

public enum SkillType
{
	None,
	Heal,
	HealAll,
	Hit,
	HitWeakest,
	HitAll,
	HitRandom,
	AddDame,
	AddArmor,
	ReducesAllDame,
	AddDameRandom,
	ReducesDameRandom,
	AddDameAll,
	ReducesRandomDame_AndHeal

}

public class Skill
{
	public SkillType skillType=SkillType.None;
	public int value1;
	public int value2;
	public int cooldown;

	public Skill()
	{
		skillType = SkillType.None;
		value1 = 0 ; 
		value2 = 0;
		cooldown = 0;
	}

}

#endregion


[System.Serializable]
public class CardInBattle 
{
	public Card card;

	public int currentHp;
	public int currentDame;
	public int currentDef;
	public int currentSpd;

	public int currentCooldown;//0 mean active

	public bool isDead= false;
	public CardInBattle(Card c)
	{
		card = new Card();
		card.atk = c.atk;
		card.defense = new Defense();
		card.defense.type = c.defense.type;
		card.defense.value1 = c.defense.value1;
		card.defense.value2 = c.defense.value2;
		card.atk = new global::Attack();
		card.atk.type = c.atk.type;
		card.atk.value = c.atk.value;
		card.atk.count = c.atk.count;
		card.hp = c.hp;
		card.id = c.id;
		card.name = c.name;
		card.spd = c.spd;
		card.skill.skillType = c.skill.skillType;
		card.skill.value1 = c.skill.value1;
		card.skill.value2 = c.skill.value2;
		card.skill.cooldown = c.skill.cooldown;
		card.rebirthTo = c.rebirthTo;
		//
		ResetValue();
	}
	void ResetValue()
	{
		currentHp = card.hp;
		currentDame = card.atk.value;
		currentDef = card.defense.value1;
		currentSpd = card.spd;
	}

	public void Attack(Team team,int pos)
	{
		for(int i = 0 ; i < card.atk.count;i++)
		{
			if(card.atk.type == AttackType.Single)
			{
				int nearest = team.GetNearestAttacked(pos);
				team.cardinfield[nearest].ReceiveDame(currentDame,this);
				if(team.IsLose() || currentHp <= 0)
					return;
			}
			else if(card.atk.type == AttackType.Splash)
			{
				int nearest = team.GetNearestAttacked(pos);
				team.cardinfield[nearest].ReceiveDame(currentDame,this);
				if(team.IsLose() || currentHp <= 0)
					return;
				
				int nearest2 = team.GetNearestAttacked(nearest-1);
				if(nearest != nearest2)
				{
					team.cardinfield[nearest2].ReceiveDame(currentDame,this);
					if(team.IsLose() || currentHp <= 0)
						return;
				}
			}
			else if(card.atk.type == AttackType.Mass)
			{
				foreach(CardInBattle c in team.cardinfield)
				{
					c.ReceiveDame(currentDame);
				}
				if(team.IsLose())
					return;
			}
		}
	}
		

	public void ReceiveDame(int damage,CardInBattle attacker = null)
	{
		if(card.defense.type == DefenseType.Deduction)
		{
			currentHp -= Mathf.Max(1,(damage-currentDef));
		}
		else  if(card.defense.type == DefenseType.Exceed)
		{
			currentHp -= Mathf.Min(card.defense.value2,damage);
		}
		else 
		{
			Mathf.Min(card.defense.value2, Mathf.Max(1,(damage-currentDef)) );
		}

		if(attacker!=null)
		{
			if(card.id == 53 || card.id ==54)
			{
				attacker.ReceiveDame(3);
			}
			else if(card.id ==55 || card.id == 56)
			{
				attacker.ReceiveDame(3);
			}
		}

		if(currentHp<=0)
		{
			if(card.rebirthTo != 0)
			{
				Card newCard = SceneController.instance.GetCardByID(card.rebirthTo);
				new CardInBattle(newCard);
				//this = new CardInBattle( 
			}
		}

	}

	public void ReceiveHP(int hp)
	{
		if(currentHp<=0)
			return;
		else
			currentHp = Mathf.Min(card.hp,currentHp+hp);
	}

	public void UpDateSkill(Team myTeam,Team enemyTeam,int pos)
	{
		if(currentCooldown<=0)
		{
			currentCooldown = card.skill.cooldown;
			switch(card.skill.skillType)
			{
			case SkillType.Heal:
				currentHp+=card.skill.value1;
				break;
			case SkillType.HealAll:
				foreach(CardInBattle c in myTeam.cardinfield)
					c.ReceiveHP(card.skill.value1);
				break;
			case SkillType.AddArmor:
				currentDef+= card.skill.value1;
				break;
			case SkillType.AddDame:
				currentDame+= card.skill.value1;
				break;
			case SkillType.Hit:
				enemyTeam.cardinfield[enemyTeam.GetNearestAttacked(pos)].ReceiveDame(card.skill.value1);
				break;
			case SkillType.HitAll:
				foreach(CardInBattle c in enemyTeam.cardinfield)
					c.ReceiveDame(card.skill.value1);
				break;
			case SkillType.HitRandom:
				enemyTeam.GetRandomCard().ReceiveDame(currentDame);
				break;
			case SkillType.HitWeakest:
				enemyTeam.GetWeakestCard().ReceiveDame(currentDame);
				break;
			case SkillType.ReducesAllDame:
				foreach(CardInBattle c in enemyTeam.cardinfield)
					c.currentDame = Math.Max(1,c.currentDame-card.skill.value1);
				break;
			case SkillType.AddDameRandom:
				myTeam.GetRandomCard().currentDame += card.skill.value1;
				break;
			case SkillType.ReducesRandomDame_AndHeal:
				CardInBattle ca1 = enemyTeam.GetRandomCard();
				ca1.currentDame = Math.Max(1,ca1.currentDame-card.skill.value1);
				ReceiveHP(card.skill.value1);
				//enemyTeam.GetRandomCard().currentDame += card.skill.value1;
				break;
			case SkillType.ReducesDameRandom:
				CardInBattle ca2 = enemyTeam.GetRandomCard();
				ca2.currentDame = Math.Max(1,ca2.currentDame-card.skill.value1);
				break;
			case SkillType.AddDameAll:
				foreach(CardInBattle c in myTeam.cardinfield)
					c.currentDame += card.skill.value1;
				break;
			}
		}
		else if(card.skill.skillType != SkillType.None)
		{
			card.skill.cooldown-=1;
		}
	}

}


public class SceneController : MonoBehaviour {
	public static SceneController instance;

	const string fileName = "Assets\\Resources\\heros.xlsx";
	int numberCard=200;

	public int loglv=5;
	public List<Card> cards;
	public Team teamA;
	public Team teamB;


	void Start () {
		instance = this;
		cards = new List<Card>();
		teamA  = new Team();
		teamB = new Team();

		FileInfo newFile = new FileInfo(fileName);
		if (newFile.Exists) 
		{
			ExcelPackage excel = new ExcelPackage(newFile);
			ExcelWorksheets sheets = excel.Workbook.Worksheets;
			LoadDefineValue(sheets["define"]);
			InitAllCards(sheets["card"]);
			Print(System.DateTime.Now.TimeOfDay.ToString(),1);
			for(int tier = 1 ; tier < 30 && cards.Count>0; tier ++)
			{
				foreach(Card c in cards)
				{
					c.chart.per = 0f;
					c.chart.all = 0;
					c.chart.win = 0;
				}
				for(int i=0;i<10000;i++)
				{
					InitTeam();
					Team teamwin = Battle(teamA,teamB);
					AddData(teamwin);
				}
				foreach(Card c in cards)
				{
					c.chart.per = ((float)c.chart.win)/c.chart.all;
				}
				List<Card> sortedList = cards.OrderByDescending(c=>(1f-c.chart.per)).ToList();
				cards = sortedList;
				int nRemove = Math.Min(4,cards.Count);
				for(int i = 0 ; i < nRemove;i++)
				{
					if(tier >16 )
						sheets["card"].Cells[cards[0].id+2,19].Value = string.Format("Tier "+(tier-16));
					else
						sheets["card"].Cells[cards[0].id+2,19].Value = "";
					sheets["card"].Cells[cards[0].id+2,16].Value = cards[0].chart.win;
					sheets["card"].Cells[cards[0].id+2,17].Value = cards[0].chart.all;

					cards.Remove(cards[0]);
				}
			}
			Print(System.DateTime.Now.TimeOfDay.ToString(),1);
			SaveData(sheets["card"]);
			excel.Save();
		}

	}
	#region excel file

	void LoadDefineValue(ExcelWorksheet sheet)
	{
		//Debug.Log("File exists" + sheet.GetValue(1,2).ToString());
		numberCard =GetInt(sheet.GetValue(1,2));

	}

	public void InitAllCards(ExcelWorksheet sheet)
	{
		cards = new List<Card>();
		for(int i=0;i<numberCard;i++)
		{
			if( sheet.GetValue(i+3,20).ToString() == "Y" )
				cards.Add(LoadCardByRow(sheet,i+3));
		}
	}
	public Card LoadCardByRow(ExcelWorksheet sheet,int row)
	{
		Card card = new Card();
		card.id = GetInt(sheet.GetValue(row,1));
		card.name = sheet.GetValue(row,2).ToString();
		//card.type = sheet.GetValue(row,2).ToString();
		//card.lv = sheet.GetValue(row,2).ToString();


		card.hp = GetInt(sheet.GetValue(row,5));
		//Debug.Log((AttackType)(Enum.Parse(typeof(AttackType),sheet.GetValue(row,6).ToString())));
		card.atk.type = (AttackType)(Enum.Parse(typeof(AttackType),sheet.GetValue(row,6).ToString()));
		card.atk.value = GetInt(sheet.GetValue(row,7));
		card.atk.count = GetInt(sheet.GetValue(row,8));

		card.defense.type = (DefenseType)(Enum.Parse(typeof(DefenseType),sheet.GetValue(row,9).ToString()));
		if(card.defense.type == DefenseType.Deduction)
		{
			card.defense.value1 = GetInt(sheet.GetValue(row,10));
		}
		else if(card.defense.type == DefenseType.Exceed)
		{
			card.defense.value2 = GetInt(sheet.GetValue(row,10));
		}
		else 
		{
			string s = sheet.GetValue(row,10).ToString();
			int.TryParse(s.Split(' ')[0],out card.defense.value1);
			int.TryParse(s.Split(' ')[1],out card.defense.value2);
		}

		card.spd = GetInt(sheet.GetValue(row,11));

		//if(Enum.TryParse(á"dsadc")); System
		try
		{
			card.skill.skillType = (SkillType)(Enum.Parse(typeof(SkillType),sheet.GetValue(row,12).ToString()));
			card.skill.value1 = GetInt(sheet.GetValue(row,13));
			card.skill.cooldown = GetInt(sheet.GetValue(row,15));

			//if(int.TryParse(sheet.GetValue(row,14).ToString(),card.skill.value2))
			//	card.skill.value2 = GetInt(sheet.GetValue(row,14));4
			int.TryParse(sheet.GetValue(row,14).ToString(),out card.skill.value2);
		}
		catch (Exception ex)
		{
			card.skill.skillType = SkillType.None;
		}

		return card;

	}
	int GetInt(object o)
	{
		return int.Parse(o.ToString());
	}
	float GetFloat(object o)
	{
		return float.Parse(o.ToString());
	}
	void Print(string s,int lv)
	{
		Debug.Log(s);
	}

	public void SaveData(ExcelWorksheet sheet)
	{
		for(int i=0;i<cards.Count;i++)
		{
			//sheet.Cells[staRow+i,winCol+3].Value=cards[i].name;
			sheet.Cells[3+i,17].Value=cards[i].chart.all;
			sheet.Cells[3+i,16].Value=cards[i].chart.win;
			cards[i].chart.per = (float)(cards[i].chart.win)/cards[i].chart.all;
		}
	}
	#endregion

	public void InitTeam()
	{

		teamA = new Team();
		for(int i = 0 ; i < 8;i++)
			teamA.cards.Add(new CardInBattle(cards[UnityEngine.Random.Range(0,cards.Count)]));

		teamB = new Team();
		for(int i = 0 ; i < 8;i++)
			teamB.cards.Add(new CardInBattle(cards[UnityEngine.Random.Range(0,cards.Count)]));



	}
	Team Battle(Team teamA,Team teamB)
	{
		teamA.MoveToBattle();
		teamB.MoveToBattle();

		int k =0;
		while(true)
		{
			k++;
			if(k>20)
				return teamA;
			for(int i =0;i<4;i++)
			{
				Team firstTeam= null;
				Team secondTeam = null;
				if(teamA.cards[i].currentSpd > teamB.cards[i].currentSpd)
				{
					firstTeam =teamA;
					secondTeam =teamB;
				}
				else
				{
					firstTeam =teamB;
					secondTeam =teamA;
				}

				//rules attack
				if(firstTeam.cardinfield.Count>i && firstTeam.cardinfield[i].currentHp > 0)
				{
					//checkskill
					firstTeam.cardinfield[i].UpDateSkill(firstTeam,secondTeam,i);
					if(secondTeam.IsLose())
						return firstTeam;
					//attack
					firstTeam.cardinfield[i].Attack(secondTeam,i);
					if(secondTeam.IsLose())
						return firstTeam;
					if(firstTeam.IsLose())
						return secondTeam;
				}



				if(secondTeam.cardinfield.Count>i && secondTeam.cardinfield[i].currentHp > 0)
				{
					//checkskill
					secondTeam.cardinfield[i].UpDateSkill(secondTeam,firstTeam,i);
					if(firstTeam.IsLose())
						return secondTeam;
					secondTeam.cardinfield[i].Attack(firstTeam,i);
					if(firstTeam.IsLose())
						return secondTeam;
					if(secondTeam.IsLose())
						return firstTeam;
				}

			}
		}

	}
	public void AddData(Team teamWin)
	{
		foreach(CardInBattle c in teamWin.cards)
		{
			GetCardByID(c.card.id).chart.win++;
		}
		foreach(CardInBattle c in teamA.cards)
		{
			GetCardByID(c.card.id).chart.all++;
		}
		foreach(CardInBattle c in teamB.cards)
		{
			GetCardByID(c.card.id).chart.all++;
		}
	}

	public Card GetCardByID(int id)
	{
		foreach(Card c in cards)
			if(c.id == id)
				return c;
		return null;
	}

}
