using System.Collections;
using System.Collections.Generic;
using UnityEngine;
using OfficeOpenXml;
using System.IO;
using System;

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
}

[System.Serializable]
public enum AttackType
{
	Single,
	Slash,
	All
}

[System.Serializable]
public enum DefenseType
{
	deduction,
	exceeded
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

#endregion


[System.Serializable]
public class CardInBattle : Card
{
	public int currentHp;
	public int currentDame;
	public bool isDead= false;
}


public class SceneController : MonoBehaviour {
	const string fileName = "Assets\\Resources\\heros.xlsx";
	int numberCard=200;

	public int loglv=5;
	public List<Card> cards;
	public Team teamA;
	public Team teamB;


	void Start () {

		Card c= new CardInBattle();
		CardInBattle cb = (CardInBattle)c;
		CardInBattle cb2 = (CardInBattle)c;

		cb.currentHp = 5;
		Debug.Log(cb2.currentHp);
		Debug.Log(cb.currentHp);
		cb2.currentHp = 7;
		Debug.Log(cb2.currentHp);
		Debug.Log(cb.currentHp);



		FileInfo newFile = new FileInfo(fileName);
		if (newFile.Exists) 
		{
			ExcelPackage excel = new ExcelPackage(newFile);
			ExcelWorksheets sheets = excel.Workbook.Worksheets;
			LoadDefineValue(sheets["define"]);
			InitAllCards(sheets["cards"]);
			Print(System.DateTime.Now.ToString(),1);
			for(int i=0;i<1000;i++)
			{
				InitTeam();
				Team teamwin = Battle(teamA,teamB);
				AddData(teamwin);
			}
			Print(System.DateTime.Now.ToString(),1);
			SaveData(sheets["heros"]);
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
	}

	public void SaveData(ExcelWorksheet sheet)
	{
		for(int i=0;i<numberCard;i++)
		{
			//sheet.Cells[staRow+i,winCol+3].Value=cards[i].name;
			sheet.Cells[3+i,15].Value=cards[i].chart.all;
			sheet.Cells[3+i,15].Value=cards[i].chart.win;
			cards[i].chart.per = (float)(cards[i].chart.win)/cards[i].chart.all;
		}
	}
	#endregion

	public void InitTeam()
	{
	}
	Team Battle(Team teamA,Team teamB)
	{
		return teamA;
	}
	public void AddData(Team teamWin)
	{
		
	}
}
