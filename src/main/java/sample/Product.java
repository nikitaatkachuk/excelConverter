package sample;

import java.util.*;

/**
 * Created by Nikita Tkachuk
 */
public class Product
{
	private String name;
	private String article;
	private String color;
	private String portionSize;
	private List<String> sizes;
	private String cost;
	private String description;
	private String imageUrl;
	private List<String> modelRow;


	public Product()
	{
	}

	public Product(String name, String article, String cost)
	{
		this.name = name;
		this.article = article;
		this.cost = cost;
	}

	public Product(String name, String article, String color, List<String> sizes, String cost, String description)
	{
		this.name = name;
		this.article = article;
		this.color = color;
		this.sizes = sizes;
		this.cost = cost;
		this.description = description;
	}

	public String getName()
	{
		return name;
	}

	public void setName(String name)
	{
		this.name = name;
	}

	public String getArticle()
	{
		return article;
	}

	public void setArticle(String article)
	{
		this.article = article;
	}

	public String getColor()
	{
		return color;
	}

	public void setColor(String color)
	{
		this.color = color;
	}

	public List<String> getSizes()
	{
		return sizes;
	}

	public void setSizes(List<String> sizes)
	{
		this.sizes = sizes;
	}

	public String getPortionSize()
	{
		return portionSize;
	}

	public void setPortionSize(String portionSize)
	{
		this.portionSize = portionSize;
	}

	public String getCost()
	{
		return cost;
	}

	public void setCost(String cost)
	{
		this.cost = cost;
	}

	public String getDescription()
	{
		return description;
	}

	public void setDescription(String description)
	{
		this.description = description;
	}

	public String getImageUrl()
	{
		return imageUrl;
	}

	public void setImageUrl(String imageUrl)
	{
		this.imageUrl = imageUrl;
	}

	public List<String> getModelRow()
	{
		return modelRow;
	}

	public void setModelRow(List<String> modelRow)
	{
		this.modelRow = modelRow;
	}

	public Map<String, String> productToExcelCellMapping()
	{
		Map<String,String> result = new LinkedHashMap<>();
		result.put("F1", this.name);
		result.put("H3", this.article);
		if(this.color != null)
		{
			result.put("H5", this.color);
		}
		StringBuilder builder = new StringBuilder();
		if(this.description != null)
		{
			builder.append(this.description);
		}
		if(this.modelRow != null && !this.modelRow.isEmpty())
		{
			builder.append("\n");
			builder.append("Постовляется модельным рядом, представленным справа");
		}
		if (builder.length() > 0)
		{
			result.put("H7", builder.toString());
		}
		if(this.sizes != null && !this.sizes.isEmpty())
		{
			int startRow = 11;
			for (int i = 0; i < this.sizes.size(); i++ )
			{
				String size = this.sizes.get(i);
				result.put("F" + startRow, size);
				if (portionSize != null)
				{
					result.put("H" + startRow, portionSize);
				}
				startRow++;
			}
		}
		result.put("AX20", this.cost);



		return result;
	}
}
