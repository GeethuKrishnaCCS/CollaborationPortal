import * as React from 'react';
import { useEffect, useState } from 'react';
import Pagination from 'react-js-pagination';
import styles from './Paging.module.scss';

export type PageUpdateCallback = (pageNumber: number) => void;

export interface IPagingProps {
  totalItems: number;
  itemsCountPerPage: number;
  onPageUpdate: PageUpdateCallback;
  currentPage: number;
}

const Paging: React.FC<IPagingProps> = (props) => {
  const [currentPage, setCurrentPage] = useState<number>(props.currentPage);

  const _pageChange = (pageNumber: number): void => {
    setCurrentPage(pageNumber);
    props.onPageUpdate(pageNumber);
  };

  useEffect(() => {
    setCurrentPage(props.currentPage);
  }, [props.currentPage]);

  return (
    <div className={styles.paginationContainer}>
      <div className={styles.searchWp__paginationContainer__pagination}>
        <Pagination
          activePage={currentPage}
          prevPageText="❮"
          nextPageText="❯"
          hideFirstLastPages={true}
          pageRangeDisplayed={0} // Hides page numbers
          itemsCountPerPage={props.itemsCountPerPage}
          totalItemsCount={props.totalItems}
          onChange={_pageChange}
          itemClass={styles.inactive}
          linkClass={styles.link}
          activeLinkClass={styles.active}
          disabledClass={styles.disabled}
        />
      </div>
    </div>
  );
};

export default Paging;
