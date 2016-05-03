# ExcelToObject #

Read excel datasheet into C# object.
![Summary image](doc/summary.png)

## Features ##
* Unity 4/5 compatible
* MIT License
* Single DLL



## Usage ##
* Table is marked with name in brackets *[Table Name]*.

#### #1) Simple case ####
* Read table into `List<T>`

![](doc/table1.png)

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
     
    public void LoadFromExcel(string filePath)
    {
    	var excelReader = new ExcelToObject.ExcelReader(filePath);
    	List<EnemyData> enemy = excelReader.ReadList<EnemyData>("EnemyData");
    }

#### #2) Array/List case ####
* A property can be a array or `List<T>`
* It consists with same column name.
* Column name can be distinguished by '#' notation. (postfix after '#' is ignored)


![](doc/table2.png)

    public class RatioData
    {
    	public int maxCount;
    	public float[] ratio;
    }
     
    List<RatioData> ratio = excelReader.ReadList<RatioData>("RatioData");

#### #3) Dictionary case ####
* A property can be a `Dictionary<TKey,TValue>`
* It consists with same column name.
* Key value is specified in '#' postfix
* Key can be any type which is convertible to.

![](doc/table3.png)

    public class StagePhaseData
    {
    	public StageType type;
    	public float interval;
    	public int maxCount;
    	public float ratio_Npc;
    	public Dictionary<EnemyType, float> ratio;
    }
     
    var stagePhase = excelReader.ReadList<StagePhaseData>("StageData");

#### #4) Result can be `Dictionary<TKey,T>` ####
* Result can be `Dictionary<TKey, TValue>`
* Specify key column name for dictionary.

![](doc/table4.png)

    public class ModData
    {
    	public float spawnInterval_Modify;
    }
     
    Dictionary<StageType, ModData> modData = excelReader.ReadDictionary<StageType, ModData>("StageModifyData", "type");


## License ##
> MIT License

Refer to [License](License) file


## Download ##

> [ExcelToObject.dll](binaries/ExcelToObject.dll)




`EOF`
