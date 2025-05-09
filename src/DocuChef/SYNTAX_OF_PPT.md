# PowerPoint ���ø� ���� �ۼ� ��ħ

## �⺻ ��Ģ

1. **DollarSignEngine ģȭ��**: ���� DollarSignEngine ��� ����� �ִ��� �����ϰ� ����
2. **���� �и�**: �����̵� ��Ʈ�� ���� ���ù���, �ؽ�Ʈ ��Ҵ� �� ���ε���
3. **������**: �������̰� ����ȭ�� �������� ��� ���Ǽ� �ش�ȭ
4. **������ ����**: PPT ������ ����� ���� �ǵ� ����

## ���� ����

### 1. �� ���ε� (�����̵� ��� ��)
```
${�Ӽ���}                   // �⺻ �Ӽ� ���ε�
${��ü.�Ӽ���}              // ��ø �Ӽ� ���ε�
${��:����������}            // ���� ������ ���
${���� ? ��1 : ��2}         // ���Ǻ� ǥ����
${�޼ҵ�()}                 // �޼ҵ� ȣ��
```

### 2. Ư�� �Լ� (�����̵� ��� ��)
```
${ppt.Image("�̹����Ӽ�")}   // �̹��� ���ε�
${ppt.Chart("��Ʈ������")}   // ��Ʈ ������ ���ε�
${ppt.Table("���̺�����")} // ǥ ������ ���ε�
```

### 3. ���� ���ù� (�����̵� ��Ʈ���� ��ġ)
```
#foreach: �÷��Ǹ� as �׸��, �ɼ�...    // �����̵� ���� (�׸�� ����� ����)
#foreach: �÷��Ǹ�, �ɼ�...             // �����̵� ���� (�׸�� �ڵ� ����)
#if: ���ǽ�, �ɼ�...                   // ���Ǻ� ó��
```

## �� ���� ����

### 1. �� ���ε� (�����̵� ��� ��)

��� �ؽ�Ʈ ��ҿ��� DollarSignEngine�� ������ �״�� ���:

```
����: ${report.title}
��¥: ${DateTime.Now:yyyy-MM-dd}
�հ�: ${items.Sum(i => i.price):C2}
����: ${status == "active" ? "Ȱ��" : "��Ȱ��"}
```

### 2. Ư�� �Լ� (�����̵� ��� ��)

#### �̹��� ���ε�
�̹��� ������ �ؽ�Ʈ��:
```
${ppt.Image("company.logo")}
${ppt.Image("product.photo", width: 300, height: 200, preserveAspectRatio: true)}
```

#### ��Ʈ ������ ���ε�
��Ʈ ������ �ؽ�Ʈ��:
```
${ppt.Chart("salesData")}
${ppt.Chart("salesData", series: "series", categories: "categories", title: "���� �Ǹŷ�")}
```

#### ǥ ������ ���ε�
ǥ ������ �ؽ�Ʈ��:
```
${ppt.Table("employeeData")}
${ppt.Table("employeeData", headers: true, startRow: 1, endRow: 10)}
```

### 3. ���� ���ù� (�����̵� ��Ʈ���� ��ġ)

#### �����̵� ���� ���ù� (#foreach)
```
#foreach: categories
#foreach: products as product
```

- �����̵��� �������� �����ϸ鼭 ������ �÷����� �� �׸� ���� �����̵带 �����մϴ�.
- �����̵� ���ο� �ε��� ����(��: `${item[0]}`, `${item[1]}` ��)�� �ִ� ���, �� �����̵�� �ش� �ε��� ������ �����ͷ� ä�����ϴ�.
- ������ �÷����� ũ�⿡ ���� �ʿ��� ��ŭ �����̵尡 �����˴ϴ�.

��: �����̵忡 `${item[0]}` ~ `${item[4]}`���� ������ �ְ� �����Ͱ� 12���� ���:
- ù ��° �����̵�: `item[0]`~`item[4]`�� �������� 0~4�� �׸����� ä����
- �� ��° �����̵�: `item[0]`~`item[4]`�� �������� 5~9�� �׸����� ä���� 
- �� ��° �����̵�: `item[0]`~`item[1]`�� �������� 10~11�� �׸����� ä������, `item[2]`~`item[4]`�� �� ������ ó����

#### ���Ǻ� ���ù�
```
#if: report.hasChart, target: "sales_chart"
#if: total > 1000, target: "warning_box", visibleWhenFalse: "success_box"
```
- `target`: ���Ǻη� ǥ��/���� ó���� ������ �̸�
- `visibleWhenFalse`: ������ ������ �� ǥ���� ��ü ������ �̸�

## ������ �ڵ����� ��Ģ

�÷��� �������� �׸� �������� �ڵ����� �����ϴ� ��Ģ:

1. **�⺻ ��Ģ**: �÷��� �̸����� ������ 's'�� ������ ���°� �׸� �������� �˴ϴ�.
   - `Items` �� `item`
   - `Products` �� `product`
   - `Categories` �� `category`

2. **���� ó��**: 's'�� ������ �ʴ� �÷��� �̸��̳� �ұ�Ģ�� �������� ���, �÷��� �̸��� �ҹ���ȭ�Ͽ� ����մϴ�.
   - `People` �� `people`
   - `Data` �� `data`

3. **����� ����**: `as` Ű���带 ����Ͽ� �������� ��������� ������ �� �ֽ��ϴ�.
   - `#foreach: Products as p`���� `${p.name}`���� ����

## ���� �÷��� ���� ó��

���� �÷����� ���ÿ� ó���ϴ� ���:

```
#foreach: Products as product, Categories as category
```

�� ���, �� �����̵忡�� `${product.name}`�� `${category.name}`���� ���� �׸� ������ �� �ֽ��ϴ�.

## ���� ���� ��ġ

1. **�� ���ε� & Ư�� �Լ�**: 
   - �ؽ�Ʈ ����, ����, ǥ, ��Ʈ �� �����̵� ����� �ؽ�Ʈ ���뿡 ��ġ
   - PowerPoint���� �ش� ��Ҹ� �����ϰ� �ؽ�Ʈ ���� ��忡�� �Է�

2. **���� ���ù�**: 
   - �����̵� ��Ʈ���� ��ġ (���� > ��Ʈ)
   - ���� ���ù��� �ʿ��� ��� ���� �� �ٿ� ��ġ

3. **��� �ĺ�**: 
   - PowerPoint���� ���� ���� �� ������ Ŭ�� �� �̸� ����
   - ������ �̸����� ���� ���ù����� ����

## �����̵� ��Ʈ �ۼ� ����

```
# �� �����̵�� ��ǰ ����� ǥ���մϴ�
#foreach: products as item
#if: products.Count > 0, target: "products_container", visibleWhenFalse: "no_products_message"
```

## ���� �ó�����

### �⺻ ���������̼� �����̵�

**�����̵� ��� ����:**
- ���� �ؽ�Ʈ ����: `${report.title}`
- ������ �ؽ�Ʈ ����: `${report.subtitle}`
- ��¥ �ؽ�Ʈ ����: `${report.date:yyyy-MM-dd}`
- �ΰ� �̹���: `${ppt.Image("company.logo")}`

**�����̵� ��Ʈ:**
```
#if: report.isConfidential, target: "confidential_watermark"
```

### ��ǰ ��� �����̵� (�迭 �ε��� ���)

**�����̵� ��� ����:**
- ���� �ؽ�Ʈ ����: `${category.name} ��ǰ ���`
- ��ǰ �׸� 1: 
  ```
  ${item[0].Id}. ${item[0].Name} - ${item[0].Description}
  ����: ${item[0].Price:C0}��
  ```
- ��ǰ �׸� 2: 
  ```
  ${item[1].Id}. ${item[1].Name} - ${item[1].Description}
  ����: ${item[1].Price:C0}��
  ```
- ��ǰ �׸� 3-5: (������ �������� ���)

**�����̵� ��Ʈ:**
```
#foreach: categories as category
  #if: category.products.Length > 0, target: "products_container"
```

### ������ ��ú��� �����̵�

**�����̵� ��� ����:**
- ���� �ؽ�Ʈ ����: `${period} �Ǹ� �м�`
- ��Ʈ ���� (�̸�: "sales_chart"): 
  ```
  ${ppt.Chart("salesData", title: "${period} �Ǹ� ����")}
  ```
- ǥ ���� (�̸�: "top_products"): 
  ```
  ${ppt.Table("topProducts", headers: true)}
  ```

**�����̵� ��Ʈ:**
```
#if: salesData.Length > 0, target: "sales_chart", visibleWhenFalse: "no_data_message"
```

### �μ��� ���� �����̵� (���� �����̵� ����)

**�����̵� ��� ����:**
- ���� �ؽ�Ʈ ����: `${dept.name} �μ�`
- �μ��� �ؽ�Ʈ ����: `�μ���: ${dept.manager}`
- �ο��� �ؽ�Ʈ ����: `�ο�: ${dept.members.Length}��`
- ���� ��Ʈ (�̸�: "dept_chart"): 
  ```
  ${ppt.Chart("dept.performance", title: "${dept.name} �μ� ����")}
  ```
- ���� 1: `${member[0].name} (${member[0].position})`
- ���� 2: `${member[1].name} (${member[1].position})`
- ���� 3-8: (������ �������� ���)

**�����̵� ��Ʈ:**
```
#foreach: departments as dept
```

### ���� �÷��� ���� ǥ�� (���� ������)

**�����̵� ��� ����:**
- ���� �ؽ�Ʈ ����: `${category.name} - ${product.name}`
- ��ǰ ����: `${product.description}`
- ī�װ� ����: `${category.description}`

**�����̵� ��Ʈ:**
```
#foreach: Products as product, Categories as category
```

�� ���������� Products�� Categories �迭�� ���ÿ� ó���ϸ�, �ε����� ���� �׸񳢸� ��Ī�˴ϴ�.