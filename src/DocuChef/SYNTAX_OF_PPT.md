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
#foreach: �÷��Ǹ� as �׸��, �ɼ�...            // �ݺ� ó�� (�׸�� ����� ����)
#foreach: �÷��Ǹ�, �ɼ�...                     // �ݺ� ó�� (�׸�� �ڵ� ����)
#if: ���ǽ�, �ɼ�...                           // ���Ǻ� ó��
#slide-foreach: �÷��Ǹ� as �׸��, �ɼ�...     // �����̵� ���� (�׸�� ����� ����)
#slide-foreach: �÷��Ǹ�, �ɼ�...              // �����̵� ���� (�׸�� �ڵ� ����)
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

#### �ݺ� ���ù�
```
#foreach: products, target: "product_box", maxItems: 6
#foreach: testimonials as testimonial, target: "quote_text"
```

- `target`: �ݺ������� ������ ����/�ؽ�Ʈ ������ �̸�
- `maxItems`: �� �����̵忡 ǥ���� �ִ� �׸� ��
- `layout`: ���̾ƿ� ���� (grid, vertical, horizontal, etc.)
- `continueOnNewSlide`: �ִ� �׸� �ʰ� �� �� �����̵忡 ��� (true/false)

#### ���Ǻ� ���ù�
```
#if: report.hasChart, target: "sales_chart"
#if: total > 1000, target: "warning_box", visibleWhenFalse: "success_box"
```
- `target`: ���Ǻη� ǥ��/���� ó���� ������ �̸�
- `visibleWhenFalse`: ������ ������ �� ǥ���� ��ü ������ �̸�

#### �����̵� ���� ���ù�
```
#slide-foreach: categories
#slide-foreach: products as product, titleTarget: "product_title", imageTarget: "product_image"
#slide-foreach: departments as dept, maxItems: 10
```
- `titleTarget`: ������ ǥ���� �ؽ�Ʈ ������ �̸�
- `imageTarget`: �̹����� ǥ���� �̹��� ������ �̸�
- `maxItems`: �����̵�� ǥ���� �ִ� �׸� �� (�� ���� �ʰ��ϸ� �� �����̵� ����)
- `maxSlides`: ������ �ִ� �����̵� ��

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
   - `#slide-foreach: Products as p`���� `${p.name}`���� ����

## ���� �÷��� ���� ó��

���� �÷����� ���ÿ� ó���ϴ� ���:

```
#slide-foreach: Products as product, Categories as category, maxItems: 5
```

�� ���, �� �����̵忡�� `${product.name}`�� `${category.name}`���� ���� �׸� ������ �� �ֽ��ϴ�.

## ��ø �÷��� ó��

��ø�� �÷����� ó���ϴ� ���:

```
#slide-foreach: Departments, maxItems: 1
  #foreach: department.Employees as emp, target: "EmployeeTemplate", maxItems: 10
```

�� ���, �� �μ��� �ϳ��� �����̵尡 �����ǰ�, �� �����̵� ������ �ش� �μ��� ���� ����� ó���˴ϴ�.

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
#foreach: products, target: "product_item", maxItems: 4, layout: "grid", continueOnNewSlide: true
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

### ��ǰ ��� �����̵�

**�����̵� ��� ����:**
- ���� �ؽ�Ʈ ����: `${category.name} ��ǰ ���`
- ��ǰ ���ø� �ؽ�Ʈ ���� (�̸�: "product_item"): 
  ```
  ${item.name}
  ����: ${item.price:C2}
  ${item.isNew ? "[����ǰ]" : ""}
  ```
- ��ǰ �̹��� ���� (�̸�: "product_image"): 
  ```
  ${ppt.Image("item.imageUrl")}
  ```

**�����̵� ��Ʈ:**
```
#slide-foreach: categories as category, maxItems: 1
  #foreach: category.products as item, target: "product_item", maxItems: 6, layout: "grid(3,2)"
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
#if: salesData.length > 0, target: "sales_chart", visibleWhenFalse: "no_data_message"
```

### �μ��� ���� �����̵� (���� �����̵� ����)

**�����̵� ��� ����:**
- ���� �ؽ�Ʈ ���� (�̸�: "dept_title"): `${dept.name} �μ�`
- �μ��� �ؽ�Ʈ ����: `�μ���: ${dept.manager}`
- �ο��� �ؽ�Ʈ ����: `�ο�: ${dept.members.length}��`
- ���� ��Ʈ (�̸�: "dept_chart"): 
  ```
  ${ppt.Chart("dept.performance", title: "${dept.name} �μ� ����")}
  ```

**�����̵� ��Ʈ:**
```
#slide-foreach: departments as dept, maxItems: 1
  #foreach: dept.members as member, target: "member_template", maxItems: 8
```

### ���� �÷��� ���� ǥ�� (���� ������)

**�����̵� ��� ����:**
- ���� �ؽ�Ʈ ����: `${category.name} - ${product.name}`
- ��ǰ ����: `${product.description}`
- ī�װ� ����: `${category.description}`

**�����̵� ��Ʈ:**
```
#slide-foreach: Products as product, Categories as category, maxItems: 1
```

�� ���������� Products�� Categories �迭�� ���ÿ� ó���ϸ�, �ε����� ���� �׸񳢸� ��Ī�˴ϴ�.