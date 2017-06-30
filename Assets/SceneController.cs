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
	public Card()
	{
		atk = new Attack();
		defense = new Defense();
		chart = new Chart();
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
	Exceed
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
	public int value;
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
	Heal,
	HealAll,
	Hit,
	HitWeakest,
	HitAll
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
	public bool isDead= false;
	public CardInBattle(Card c)
	{
		card = new Card();
		card.atk = c.atk;
		card.defense = new Defense();
		card.defense.type = c.defense.type;
		card.defense.value = c.defense.value;
		card.atk = new global::Attack();
		card.atk.type = c.atk.type;
		card.atk.value = c.atk.value;
		card.atk.count = c.atk.count;
		card.hp = c.hp;
		card.id = c.id;
		card.name = c.name;
		card.spd = c.spd;

		//
		ResetValue();
	}
	void ResetValue()
	{
		currentHp = card.hp;
		currentDame = card.atk.value;
		currentDef = card.defense.value;
		currentSpd = card.spd;
	}

	public void Attack(Team team,int pos)
	{
		for(int i = 0 ; i < card.atk.count;i++)
		{
			if(card.atk.type == AttackType.Single)
			{
				int nearest = team.GetNearestAttacked(pos);
				team.cardinfield[nearest].ReceiveDame(currentDame);// .currentHp -= Mathf.Max(0,(currentDame-team.cardinfield[nearest].currentDef));
				if(team.IsLose())
					return;
			}
			else if(card.atk.type == AttackType.Splash)
			{
				int nearest = team.GetNearestAttacked(pos);
				team.cardinfield[nearest].ReceiveDame(currentDame);//.currentHp -= Mathf.Max(0,(currentDame-team.cardinfield[nearest].currentDef));
				if(team.IsLose())
					return;
				
				int nearest2 = team.GetNearestAttacked(nearest-1);
				if(nearest != nearest2)
				{
					team.cardinfield[nearest2].ReceiveDame(currentDame);// .currentHp -= Mathf.Max(0,(currentDame-team.cardinfield[nearest2].currentDef));
					if(team.IsLose())
						return;
				}
			}
			else if(card.atk.type == AttackType.Mass)
			{
				foreach(CardInBattle c in team.cardinfield)
				{
					c.ReceiveDame(currentDame);// .currentHp -= Mathf.Max(0,(currentDame-c.currentDef));
				}
				if(team.IsLose())
					return;
			}
		}
	}

	public void ReceiveDame(int damage)
	{
		if(card.defense.type == DefenseType.Deduction)
		{
			currentHp -= Mathf.Max(0,(damage-currentDef));
		}
		else 
		{
			currentHp -= Mathf.Min(currentDef,damage);
		}
	}

}


public class SceneController : MonoBehaviour {
	const string fileName = "Assets\\Resources\\heros.xlsx";
	int numberCard=200;

	public int loglv=5;
	public List<Card> cards;
	public Team teamA;
	public Team teamB;


	void Start () {
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
			for(int tier = 1 ; tier < 10 && cards.Count>0; tier ++)
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
				int nRemove = Math.Min(10,cards.Count);
				for(int i = 0 ; i < nRemove;i++)
				{
					if(tier >5 )
						sheets["card"].Cells[cards[0].id+2,19].Value = string.Format("Tier "+(tier-5));
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
		card.defense.value = GetInt(sheet.GetValue(row,10));

		card.spd = GetInt(sheet.GetValue(row,11));
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
					firstTeam.cardinfield[i].Attack(secondTeam,i);
					if(secondTeam.IsLose())
						return firstTeam;
				}

				if(secondTeam.cardinfield.Count>i && secondTeam.cardinfield[i].currentHp > 0)
				{
					secondTeam.cardinfield[i].Attack(firstTeam,i);
					if(firstTeam.IsLose())
						return secondTeam;
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
